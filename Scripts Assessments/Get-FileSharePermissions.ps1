<#
.SYNOPSIS
    v2.2 — Exporte les permissions NTFS d'un FileShare avec scan haute performance (40 TB / 4,2M dossiers).

.DESCRIPTION
    Ce script analyse un ou plusieurs chemins locaux/UNC et exporte les permissions NTFS
    des dossiers rencontrés.

    Deux modes d'utilisation :
    - Mode Single  : -CheminUNC <chemin> (rétrocompatible)
    - Mode Mapping : -MappingCsv <chemin du fichier FileShareMapping.csv>

    Optimisations v2 (cumulées) :
    1. Skip inherited-only   — les dossiers n'ayant que des ACE héritées sont ignorés.
    2. Parallélisation PS7+  — ForEach-Object -Parallel (ThrottleLimit = 8 par défaut).
    3. API .NET directe      — DirectoryInfo + FileSystemAclExtensions (méthode d'extension PS7+).
    4. Streaming CSV         — StreamWriter au fil de l'eau -> RAM stable < 100 MB.
    5. Fallback Get-Acl      — si l'API .NET échoue, bascule automatique sur Get-Acl (100% portable).

#>

[CmdletBinding(DefaultParameterSetName = 'Single')]
param(
    [Parameter(Mandatory, ParameterSetName = 'Single')]
    [ValidateNotNullOrEmpty()]
    [string]$CheminUNC,

    [Parameter(Mandatory, ParameterSetName = 'Mapping')]
    [ValidateNotNullOrEmpty()]
    [string]$MappingCsv,

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$OutputPath = (Join-Path -Path $PSScriptRoot -ChildPath "Output"),

    [Parameter(Mandatory = $false)]
    [ValidateRange(0, 100)]
    [int]$Depth = 0,

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 32)]
    [int]$ThrottleLimit = 8,

    [Parameter(Mandatory = $false)]
    [switch]$IncludeInheritedOnly,

    [Parameter(Mandatory = $false)]
    [ValidateRange(100, 10000)]
    [int]$BatchSize = 1000,

    [Parameter(Mandatory = $false)]
    [PSCustomObject]$Run = $null,

    [Parameter(Mandatory = $false)]
    [System.Management.Automation.PSCredential]
    [System.Management.Automation.Credential()]
    $Credential,

    [Parameter(Mandatory = $false)]
    [string]$Server = 'ntx-pa-fs01.primagaz.fr'
)

begin {
    # Chargement de l'assembly DANS le thread principal (utile pour le mode séquentiel)
    try { Add-Type -AssemblyName System.IO.FileSystem.AccessControl -ErrorAction SilentlyContinue } catch {}

    Import-Module (Join-Path -Path $PSScriptRoot -ChildPath "Modules\PrimaGAZ.Assessment.psm1") -Force
    $outputModulePath = Join-Path -Path $PSScriptRoot -ChildPath "Modules\PrimaGAZ.Output.psm1"
    if (Test-Path $outputModulePath) { Import-Module $outputModulePath -Force }
    $smbModulePath = Join-Path -Path $PSScriptRoot -ChildPath "Modules\PrimaGAZ.SmbCredential.psm1"
    if (Test-Path $smbModulePath) {
        Import-Module $smbModulePath -Force -ErrorAction $(if ($Credential) { 'Stop' } else { 'SilentlyContinue' })
    } elseif ($Credential) {
        throw "Module PrimaGAZ.SmbCredential introuvable. Chemin attendu : $smbModulePath"
    }
    $script:smbConnected = $false

    $ScriptName     = Split-Path -Leaf $PSCommandPath
    $ScriptBaseName = [System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)
    $StartTime      = Get-Date
    $WarningCount   = 0
    $ErrorCount     = 0
    $ItemsScanned   = 0
    $Status         = "Success"

    $CSV_HEADER = '"CheminDossier";"Proprietaire";"Identite";"TypeAcces";"Droits";"EstHerite";"InheritanceFlags";"PropagationFlags";"DateAnalyse";"ResultatAnalyse";"NomFileShare"'

    # =========================================================
    # Fonctions utilitaires
    # =========================================================

    function Get-RelativeDepth {
        param([string]$RootPath, [string]$CurrentPath)
        $normalizedRoot    = $RootPath.TrimEnd([char]'\')
        $normalizedCurrent = $CurrentPath.TrimEnd([char]'\')
        if ($normalizedCurrent.Length -le $normalizedRoot.Length) { return 0 }
        $relativePath = $normalizedCurrent.Substring($normalizedRoot.Length).TrimStart([char]'\')
        if ([string]::IsNullOrWhiteSpace($relativePath)) { return 0 }
        return ($relativePath -split '\\').Count
    }

    function Get-SafeFileName {
        param([string]$Value)
        if ([string]::IsNullOrWhiteSpace($Value)) { return "Inconnu" }
        $safe = $Value
        foreach ($invalidChar in [System.IO.Path]::GetInvalidFileNameChars()) {
            $safe = $safe.Replace([string]$invalidChar, "_")
        }
        return $safe.Trim()
    }

    function Get-NomFileShareFromPath {
        param([string]$Path)
        if ([string]::IsNullOrWhiteSpace($Path)) { return "FileShare" }
        $trimmed  = $Path.TrimEnd([char]'\', [char]'/')
        if ([string]::IsNullOrWhiteSpace($trimmed)) { return "FileShare" }
        $parts    = $trimmed -split '[/\\]'
        $nonEmpty = @($parts | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        if ($nonEmpty.Count -eq 0) { return "FileShare" }
        return $nonEmpty[-1]
    }

    function Get-ExecutionAccount {
        try {
            if ($PSVersionTable.PSEdition -eq 'Desktop') {
                return [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
            }
        } catch { }
        try { return (whoami) } catch { }
        if (-not [string]::IsNullOrWhiteSpace($env:USERDOMAIN) -and -not [string]::IsNullOrWhiteSpace($env:USERNAME)) {
            return "$($env:USERDOMAIN)\$($env:USERNAME)"
        }
        if (-not [string]::IsNullOrWhiteSpace($env:USERNAME)) { return $env:USERNAME }
        return "Inconnu"
    }

    # =========================================================
    # *** CORRECTIF v2.2 — Helper ACL multi-versions ***
    # =========================================================
  
    function Get-DirectoryAclSafe {
        param([Parameter(Mandatory)][string]$Path)

        # Méthode 1 : extension method .NET Core/PS7+
        try {
            $di = [System.IO.DirectoryInfo]::new($Path)
            return $di.GetAccessControl()
        } catch { }

        # Méthode 2 : méthode d'instance .NET Framework (PS5.1)
        try {
            $di = [System.IO.DirectoryInfo]::new($Path)
            return $di.GetAccessControl()
        } catch { }

        # Méthode 3 : fallback ultime — Get-Acl
        return Get-Acl -LiteralPath $Path -ErrorAction Stop
    }

    # =========================================================
    # Scan ACL d'un N1 — cœur de la logique v2
    # =========================================================

    function Invoke-N1PermissionsScan {
        param(
            [Parameter(Mandatory = $true)][string]$CheminN1,
            [Parameter(Mandatory = $true)][string]$NomFS,
            [Parameter(Mandatory = $true)][int]$Depth,
            [Parameter(Mandatory = $true)][string]$CsvPath,
            [Parameter(Mandatory = $true)][bool]$DoSkipInherited,
            [Parameter(Mandatory = $true)][int]$ThrottleLimit,
            [Parameter(Mandatory = $true)][int]$BatchSize
        )

        $errorCount    = 0
        $skippedCount  = 0
        $brokenCount   = 0
        $aceCount      = 0
        $writeCounter  = 0

        $scanErrorsList = New-Object 'System.Collections.Generic.List[object]'
        $brokenList     = New-Object 'System.Collections.Generic.List[string]'
        $safeNomFS      = Get-SafeFileName -Value $NomFS

        $directories = @(Invoke-SafeRecursiveScan -RootPath $CheminN1 -DirectoriesOnly `
            -ErrorCollection $scanErrorsList -NomFileShare $NomFS `
            -ProgressActivity "Enumération des dossiers pour ACL")

        try {
            $rootItem    = [System.IO.DirectoryInfo]::new($CheminN1)
            $directories = @($rootItem) + $directories
        } catch {
            Write-Log "Impossible d'accéder au dossier racine : $CheminN1 - $($_.Exception.Message)" "WARN"
        }

        if ($Depth -gt 0) {
            $directories = $directories | Where-Object {
                (Get-RelativeDepth -RootPath $CheminN1 -CurrentPath $_.FullName) -le $Depth
            }
        }

        $itemsScannedLocal = ($directories | Measure-Object).Count
        Write-Log "Dossiers à analyser pour '$NomFS' : $itemsScannedLocal" "INFO"

        $writer = [System.IO.StreamWriter]::new($CsvPath, $false, [System.Text.UTF8Encoding]::new($true))

        try {
            $writer.WriteLine($script:CSV_HEADER)

            if ($PSVersionTable.PSVersion.Major -ge 7 -and $ThrottleLimit -gt 1 -and $itemsScannedLocal -gt 0) {
                # =================================================
                # MODE PARALLÈLE PS7+
                # =================================================
                Write-Log "Mode parallèle PS7+ activé (ThrottleLimit=$ThrottleLimit)" "INFO"

                $progressIndex = 0

                $directories | ForEach-Object -ThrottleLimit $ThrottleLimit -Parallel {
                    # Chargement de l'assembly dans chaque runspace isolé
                    try { Add-Type -AssemblyName System.IO.FileSystem.AccessControl -ErrorAction SilentlyContinue } catch {}

                    $dir          = $_
                    $nomFsLocal   = $using:NomFS
                    $doSkip       = $using:DoSkipInherited

                    # *** Helper inliné — les fonctions ne traversent pas les runspaces ***
                    $getAcl = {
                        param($p)
                        # Méthode 1 : extension PS7+
                        try {
                            $di = [System.IO.DirectoryInfo]::new($p)
                            return [System.IO.FileSystemAclExtensions]::GetAccessControl($di)
                        } catch {}
                        # Méthode 2 : instance PS5.1
                        try {
                            $di = [System.IO.DirectoryInfo]::new($p)
                            return $di.GetAccessControl()
                        } catch {}
                        # Méthode 3 : fallback Get-Acl
                        return Get-Acl -LiteralPath $p -ErrorAction Stop
                    }

                    try {
                        $acl = & $getAcl $dir.FullName

                        $hasExplicitAce = $false
                        foreach ($ace in $acl.Access) {
                            if (-not $ace.IsInherited) { $hasExplicitAce = $true; break }
                        }
                        $isBroken = $acl.AreAccessRulesProtected

                        if ($doSkip -and -not $hasExplicitAce -and -not $isBroken) {
                            return [PSCustomObject]@{
                                Type       = 'skip'
                                Lines      = $null
                                BrokenPath = $null
                                Error      = $null
                            }
                        }

                        $dateAnalyse = [datetime]::Now.ToString('yyyy-MM-dd HH:mm:ss')
                        $pathEsc     = $dir.FullName -replace '"', '""'
                        $ownerEsc    = ($acl.Owner -as [string]) -replace '"', '""'
                        $nomFsEsc    = $nomFsLocal -replace '"', '""'

                        $lines = [System.Collections.Generic.List[string]]::new()
                        foreach ($access in $acl.Access) {
                            $lines.Add(
                                '"{0}";"{1}";"{2}";"{3}";"{4}";"{5}";"{6}";"{7}";"{8}";"{9}";"{10}"' -f
                                $pathEsc,
                                $ownerEsc,
                                ($access.IdentityReference.ToString() -replace '"', '""'),
                                $access.AccessControlType.ToString(),
                                ($access.FileSystemRights.ToString() -replace '"', '""'),
                                $access.IsInherited,
                                $access.InheritanceFlags.ToString(),
                                $access.PropagationFlags.ToString(),
                                $dateAnalyse,
                                'Permission exportée',
                                $nomFsEsc
                            )
                        }

                        return [PSCustomObject]@{
                            Type       = if ($isBroken) { 'broken' } else { 'acl' }
                            Lines      = $lines
                            BrokenPath = if ($isBroken) { $dir.FullName } else { $null }
                            Error      = $null
                        }
                    } catch {
                        return [PSCustomObject]@{
                            Type       = 'error'
                            Lines      = $null
                            BrokenPath = $null
                            Error      = [PSCustomObject]@{
                                Chemin        = $dir.FullName
                                TypeErreur    = if ($_.Exception.Message -match 'denied|refus|Unauthorized|non autoris') { 'AccessDenied' } else { 'AclError' }
                                ExceptionType = $_.Exception.GetType().Name
                                MessageErreur = $_.Exception.Message
                            }
                        }
                    }

                } | ForEach-Object {
                    $result = $_
                    $progressIndex++

                    if ($progressIndex % 500 -eq 0 -or $progressIndex -eq $itemsScannedLocal) {
                        $pct = [math]::Min(100, [math]::Round(($progressIndex / [math]::Max(1, $itemsScannedLocal)) * 100))
                        Write-Progress -Activity "Analyse des permissions NTFS" -Status "[$NomFS] $progressIndex / $itemsScannedLocal" -PercentComplete $pct
                    }

                    switch ($result.Type) {
                        'skip' { $skippedCount++ }
                        { $_ -in 'acl', 'broken' } {
                            foreach ($line in $result.Lines) {
                                $writer.WriteLine($line)
                            }
                            $aceCount     += $result.Lines.Count
                            $writeCounter++
                            if ($writeCounter % $BatchSize -eq 0) { $writer.Flush() }

                            if ($result.Type -eq 'broken') {
                                $brokenCount++
                                $brokenList.Add($result.BrokenPath)
                            }
                        }
                        'error' {
                            $errorCount++
                            $scanErrorsList.Add([PSCustomObject]@{
                                NomFileShare  = $NomFS
                                Chemin        = $result.Error.Chemin
                                TypeErreur    = $result.Error.TypeErreur
                                ExceptionType = $result.Error.ExceptionType
                                MessageErreur = $result.Error.MessageErreur
                                DateDetection = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
                            })
                        }
                    }

                    # *** Sanity check renforcé : détection précoce dès 50 erreurs consécutives ***
                    if ($progressIndex -eq 50 -and $aceCount -eq 0 -and $errorCount -ge 40) {
                        Write-Log "[ALERTE] Sanity check précoce : $errorCount erreurs sur les 50 premiers dossiers, 0 ACE exportée. Probable bug runspace ou ACL inaccessible. Recommandation : -ThrottleLimit 1 pour forcer le mode séquentiel avec fallback Get-Acl." "WARN"
                    }
                }

                if ($aceCount -eq 0 -and $skippedCount -ge 100) {
                    Write-Log "[WARN] Sanity check: $skippedCount dossiers analysés en parallèle sans aucune ACE explicite. Possible bug runspace. Recommandation : relancer avec -ThrottleLimit 1." "WARN"
                }

            } else {
                # =================================================
                # MODE SÉQUENTIEL — PS5.1 ou ThrottleLimit = 1
                # =================================================
                if ($PSVersionTable.PSVersion.Major -lt 7) {
                    Write-Log "PowerShell 5.1 détecté — fallback séquentiel. Recommandation : passer à PowerShell 7 pour x8 en performance." "WARN"
                } else {
                    Write-Log "Mode séquentiel (ThrottleLimit=1) — utilisation de Get-DirectoryAclSafe avec fallback Get-Acl." "INFO"
                }

                $index = 0
                foreach ($dir in $directories) {
                    $index++
                    if ($index % 100 -eq 0 -or $index -eq $itemsScannedLocal) {
                        $pct = [math]::Min(100, [math]::Round(($index / [math]::Max(1, $itemsScannedLocal)) * 100))
                        Write-Progress -Activity "Analyse des permissions NTFS" -Status "[$NomFS] $index / $itemsScannedLocal" -PercentComplete $pct
                    }

                    Write-Verbose "Analyse ACL [$index/$itemsScannedLocal] : $($dir.FullName)"

                    try {
                        # *** CORRECTIF : utilisation du helper multi-versions au lieu de [System.IO.Directory]::GetAccessControl ***
                        $acl = Get-DirectoryAclSafe -Path $dir.FullName

                        $hasExplicitAce = $false
                        foreach ($ace in $acl.Access) {
                            if (-not $ace.IsInherited) { $hasExplicitAce = $true; break }
                        }
                        $isBroken = $acl.AreAccessRulesProtected

                        if ($DoSkipInherited -and -not $hasExplicitAce -and -not $isBroken) {
                            $skippedCount++
                            continue
                        }

                        $dateAnalyse = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
                        $pathEsc     = $dir.FullName -replace '"', '""'
                        $ownerEsc    = ($acl.Owner -as [string]) -replace '"', '""'
                        $nomFsEsc    = $NomFS -replace '"', '""'

                        foreach ($access in $acl.Access) {
                            $line = '"{0}";"{1}";"{2}";"{3}";"{4}";"{5}";"{6}";"{7}";"{8}";"{9}";"{10}"' -f
                                $pathEsc,
                                $ownerEsc,
                                ($access.IdentityReference.ToString() -replace '"', '""'),
                                $access.AccessControlType.ToString(),
                                ($access.FileSystemRights.ToString() -replace '"', '""'),
                                $access.IsInherited,
                                $access.InheritanceFlags.ToString(),
                                $access.PropagationFlags.ToString(),
                                $dateAnalyse,
                                'Permission exportée',
                                $nomFsEsc
                            $writer.WriteLine($line)
                            $aceCount++
                        }

                        $writeCounter++
                        if ($writeCounter % $BatchSize -eq 0) { $writer.Flush() }

                        if ($isBroken) {
                            $brokenCount++
                            $brokenList.Add($dir.FullName)
                        }

                    } catch {
                        $errorCount++
                        $scanErrorsList.Add([PSCustomObject]@{
                            NomFileShare  = $NomFS
                            Chemin        = $dir.FullName
                            TypeErreur    = if ($_.Exception.Message -match 'denied|refus|Unauthorized|non autoris') { 'AccessDenied' } else { 'AclError' }
                            ExceptionType = $_.Exception.GetType().Name
                            MessageErreur = $_.Exception.Message
                            DateDetection = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
                        })

                        # *** Sanity check séquentiel : détection précoce ***
                        if ($index -eq 50 -and $aceCount -eq 0 -and $errorCount -ge 40) {
                            Write-Log "[ALERTE] $errorCount erreurs sur les 50 premiers dossiers, 0 ACE exportée. Vérifier les droits du compte d'exécution." "WARN"
                        }
                    }
                }
            }

            if ($aceCount -eq 0) {
                $dateVide = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
                $writer.WriteLine(
                    '"";"";"";"";"";"";"";"";"' + $dateVide + '";"Aucune permission exportée";"' + ($NomFS -replace '"', '""') + '"'
                )
            }

            $writer.Flush()

        } finally {
            $writer.Close()
        }

        # --- Reporting des erreurs ---
        if ($scanErrorsList.Count -gt 0) {
            $denied = @($scanErrorsList | Where-Object { $_.TypeErreur -eq 'AccessDenied' })
            $others = @($scanErrorsList | Where-Object { $_.TypeErreur -ne 'AccessDenied' })
            $errBase   = if ($script:RunErrorsPath) { $script:RunErrorsPath } else { $script:OutputPath }
            $errSuffix = if ($script:RunMode)       { "" }                  else { "_${script:timestamp}" }

            if ($denied.Count -gt 0) {
                Write-Log "Sous-dossiers refusés sur $NomFS : $($denied.Count)" "WARN"
                $deniedCsv = Join-Path $errBase "AccessDenied_${safeNomFS}${errSuffix}.csv"
                $denied | Export-Csv -Path $deniedCsv -NoTypeInformation -Delimiter ';' -Encoding UTF8
                Write-Log "CSV des refus : $deniedCsv" "WARN"
                $denied | Select-Object -First 5 | ForEach-Object {
                    Write-Log "  -> Refusé : $($_.Chemin) [$($_.ExceptionType)]" "WARN"
                }
                if ($denied.Count -gt 5) {
                    Write-Log "  -> ... et $($denied.Count - 5) autre(s) refus (voir CSV)" "WARN"
                }
                $script:WarningCount += $denied.Count
            }

            if ($others.Count -gt 0) {
                Write-Log "Autres erreurs sur $NomFS : $($others.Count)" "WARN"
                $errCsv = Join-Path $errBase "ScanErrors_${safeNomFS}${errSuffix}.csv"
                $others | Export-Csv -Path $errCsv -NoTypeInformation -Delimiter ';' -Encoding UTF8
                Write-Log "CSV des erreurs : $errCsv" "WARN"
                $others | Select-Object -First 3 | ForEach-Object {
                    Write-Log "  -> $($_.ExceptionType) : $($_.Chemin)" "WARN"
                    Write-Log "     Message: $($_.MessageErreur)" "WARN"
                }
                $script:WarningCount += $others.Count
            }
        }

        return @{
            ItemsScanned             = $itemsScannedLocal
            AclErrors                = $errorCount
            SkippedInherited         = $skippedCount
            InheritanceBrokenCount   = $brokenCount
            AceExported              = $aceCount
            InheritanceBrokenPaths   = $brokenList
            HasResults               = ($aceCount -gt 0)
        }
    }

    # =========================================================
    # Initialisation
    # =========================================================

    $script:RunErrorsPath = $null
    $script:RunMode       = $false
    Write-Log "Initialisation de l'analyse des permissions NTFS"
    Initialize-OutputPath -Path $OutputPath

    $timestamp           = Get-Date -Format "yyyyMMdd_HHmmss"
    $csvConsolidePath    = Join-Path -Path $OutputPath -ChildPath "Permissions_NTFS_${timestamp}.csv"
    $metadataPath        = Join-Path -Path $OutputPath -ChildPath "Permissions_NTFS_Metadata_${timestamp}.json"
    $brokenCsvPath       = Join-Path -Path $OutputPath -ChildPath "Permissions_InheritanceBroken_${timestamp}.csv"
    $LogPath             = Join-Path -Path $OutputPath -ChildPath "$ScriptBaseName`_$timestamp.log"
    Set-LogFile -Path $LogPath

    if ($null -ne $Run) {
        $script:RunErrorsPath = $Run.Errors
        $script:RunMode       = $true
        $OutputPath           = $Run.Csv
        $metadataPath         = Join-Path $Run.Metadata "Permissions_NTFS.json"
        $LogPath              = Join-Path $Run.Logs "Get-FileSharePermissions.log"
        $csvConsolidePath     = Join-Path $Run.Csv "Permissions_NTFS.csv"
        $brokenCsvPath        = Join-Path $Run.Csv "Permissions_InheritanceBroken.csv"
        Set-LogFile -Path $LogPath
    }

    Write-Log "=== Démarrage $ScriptName (v2.2) ===" "INFO"
    Write-Log "VM source : $env:COMPUTERNAME" "INFO"
    Write-Log "Compte d'exécution : $(Get-ExecutionAccount)" "INFO"
    Write-Log "OutputPath : $OutputPath" "INFO"
    Write-Log "Mode skip inherited-only : $(-not $IncludeInheritedOnly)" "INFO"
    Write-Log "ThrottleLimit : $ThrottleLimit (PS$($PSVersionTable.PSVersion.Major))" "INFO"
    Write-Log "BatchSize : $BatchSize" "INFO"

    # Test rapide de l'API ACL au démarrage
    try {
        $testPath = if ($PSCmdlet.ParameterSetName -eq 'Single') { $CheminUNC } else { $env:TEMP }
        if (Test-Path $testPath) {
            $null = Get-DirectoryAclSafe -Path $testPath
            Write-Log "Test ACL API : OK sur $testPath" "INFO"
        }
    } catch {
        Write-Log "Test ACL API : ÉCHEC sur le chemin de test - $($_.Exception.Message)" "WARN"
    }

    $cheminsAScanner = New-Object 'System.Collections.Generic.List[object]'

    if ($PSCmdlet.ParameterSetName -eq 'Mapping') {
        if (-not (Test-Path -Path $MappingCsv)) {
            throw "Le fichier mapping '$MappingCsv' est introuvable."
        }

        $mappingData = @(Import-Csv -Path $MappingCsv -Delimiter ';' -Encoding UTF8)
        if ($mappingData.Count -eq 0) {
            throw "Le fichier mapping '$MappingCsv' est vide."
        }

        $columns = @($mappingData[0].PSObject.Properties.Name)
        if (-not ($columns -contains 'CheminUNC')) {
            throw "La colonne obligatoire 'CheminUNC' est absente du mapping CSV."
        }
        if (-not ($columns -contains 'NomFileShare')) {
            throw "La colonne obligatoire 'NomFileShare' est absente du mapping CSV."
        }

        foreach ($row in $mappingData) {
            if ([string]::IsNullOrWhiteSpace($row.CheminUNC) -or [string]::IsNullOrWhiteSpace($row.NomFileShare)) {
                Write-Log "Ligne mapping ignorée (CheminUNC ou NomFileShare vide)." "WARN"
                $WarningCount++
                continue
            }
            $cheminsAScanner.Add([PSCustomObject]@{
                CheminUNC    = $row.CheminUNC.Trim()
                NomFileShare = $row.NomFileShare.Trim()
            })
        }

        if ($cheminsAScanner.Count -eq 0) {
            throw "Aucune entrée exploitable trouvée dans '$MappingCsv'."
        }
    } else {
        Test-SourcePath -Path $CheminUNC
        $cheminsAScanner.Add([PSCustomObject]@{
            CheminUNC    = $CheminUNC
            NomFileShare = Get-NomFileShareFromPath -Path $CheminUNC
        })
    }

    $globalStats = @{
        N1Reussis              = 0
        N1Refuses              = 0
        N1Erreurs              = 0
        DossiersEnumeres       = 0
        DossiersSkipped        = 0
        DossiersInhBroken      = 0
        AceExportees           = 0
        ErreursACL             = 0
    }

    $csvParN1Paths = New-Object 'System.Collections.Generic.List[string]'
    $allBrokenPaths = New-Object 'System.Collections.Generic.List[string]'
}

process {
    if ($Credential) {
        $smbResult = Connect-PrimaGazFileShare -Server $Server -Credential $Credential
        Test-FileShareIdentity -Server $Server -ExpectedUserName $Credential.UserName
        $script:smbConnected = -not $smbResult.AlreadyConnected
    }

    Write-Log "Mode d'exécution : $($PSCmdlet.ParameterSetName)"

    $isSingleMode = ($PSCmdlet.ParameterSetName -eq 'Single')

    foreach ($entry in $cheminsAScanner) {
        $cheminN1  = $entry.CheminUNC
        $nomFS     = $entry.NomFileShare
        $safeNomFS = Get-SafeFileName -Value $nomFS

        Write-Log "=== Démarrage scan ACL : $nomFS ($cheminN1) ===" "INFO"

        try {
            $null = Get-Item -Path $cheminN1 -ErrorAction Stop
        } catch [System.UnauthorizedAccessException] {
            Write-Log "ACCESS DENIED sur $nomFS - skip (compte sans droits sur ce share)" "WARN"
            $WarningCount++
            $globalStats.N1Refuses++
            continue
        } catch {
            Write-Log "Chemin inaccessible : $nomFS - $($_.Exception.Message)" "ERROR"
            $ErrorCount++
            $globalStats.N1Erreurs++
            continue
        }

        if ($isSingleMode) {
            $csvN1Path = $csvConsolidePath
        } else {
            $csvN1Path = Join-Path -Path $OutputPath -ChildPath "Permissions_NTFS_${safeNomFS}_${timestamp}.csv"
        }

        try {
            $scanN1 = Invoke-N1PermissionsScan `
                -CheminN1        $cheminN1 `
                -NomFS           $nomFS `
                -Depth           $Depth `
                -CsvPath         $csvN1Path `
                -DoSkipInherited (-not $IncludeInheritedOnly) `
                -ThrottleLimit   $ThrottleLimit `
                -BatchSize       $BatchSize

            $ItemsScanned += [int]$scanN1.ItemsScanned

            $globalStats.DossiersEnumeres  += [int]$scanN1.ItemsScanned
            $globalStats.DossiersSkipped   += [int]$scanN1.SkippedInherited
            $globalStats.DossiersInhBroken += [int]$scanN1.InheritanceBrokenCount
            $globalStats.AceExportees      += [int]$scanN1.AceExported
            $globalStats.ErreursACL        += [int]$scanN1.AclErrors

            if ($scanN1.InheritanceBrokenPaths.Count -gt 0) {
                $allBrokenPaths.AddRange($scanN1.InheritanceBrokenPaths)
            }

            if (-not $isSingleMode) {
                $csvParN1Paths.Add($csvN1Path)
                Write-Log "CSV N1 généré : $csvN1Path" "SUCCESS"
            }

            if ($scanN1.AclErrors -gt 0) {
                Write-Log "Erreurs ACL non bloquantes sur $nomFS : $($scanN1.AclErrors)" "WARN"
                $WarningCount++
            }

            $globalStats.N1Reussis++

        } catch [System.UnauthorizedAccessException] {
            Write-Log "Access denied en cours de scan sur $nomFS - skip partiel" "WARN"
            $WarningCount++
            $globalStats.N1Refuses++
            continue
        } catch {
            Write-Log "Erreur scan ACL $nomFS : $($_.Exception.Message)" "ERROR"
            $ErrorCount++
            $globalStats.N1Erreurs++
            continue
        }
    }
}

end {
    if ($script:smbConnected) {
        try { Disconnect-PrimaGazFileShare -Server $Server } catch { Write-Verbose "Disconnect-PrimaGazFileShare: $($_.Exception.Message)" }
        $script:smbConnected = $false
    }
    try {
        Write-Progress -Activity "Analyse des permissions NTFS" -Completed

        if ($PSCmdlet.ParameterSetName -eq 'Mapping' -and $csvParN1Paths.Count -gt 0) {
            Write-Log "Fusion des CSV N1 vers le consolidé : $csvConsolidePath" "INFO"
            $mergeWriter = [System.IO.StreamWriter]::new($csvConsolidePath, $false, [System.Text.UTF8Encoding]::new($true))
            try {
                $mergeWriter.WriteLine($CSV_HEADER)
                foreach ($csvN1 in $csvParN1Paths) {
                    if (-not (Test-Path $csvN1)) { continue }
                    $reader = [System.IO.StreamReader]::new($csvN1, [System.Text.UTF8Encoding]::new($true))
                    try {
                        $null = $reader.ReadLine()
                        while (-not $reader.EndOfStream) {
                            $mergeWriter.WriteLine($reader.ReadLine())
                        }
                    } finally {
                        $reader.Close()
                    }
                }
                $mergeWriter.Flush()
            } finally {
                $mergeWriter.Close()
            }
            Write-Log "CSV consolidé généré : $csvConsolidePath" "SUCCESS"
        } elseif ($PSCmdlet.ParameterSetName -eq 'Single') {
            Write-Log "CSV consolidé (mode Single) : $csvConsolidePath" "SUCCESS"
        } else {
            $writer = [System.IO.StreamWriter]::new($csvConsolidePath, $false, [System.Text.UTF8Encoding]::new($true))
            try {
                $writer.WriteLine($CSV_HEADER)
                $dateVide = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
                $writer.WriteLine('"";"";"";"";"";"";"";"";"' + $dateVide + '";"Aucune permission exportée";"[CONSOLIDE]"')
                $writer.Flush()
            } finally {
                $writer.Close()
            }
        }

        if ($allBrokenPaths.Count -gt 0) {
            Write-Log "Génération du CSV héritage cassé : $brokenCsvPath ($($allBrokenPaths.Count) dossiers)" "INFO"
            $brokenWriter = [System.IO.StreamWriter]::new($brokenCsvPath, $false, [System.Text.UTF8Encoding]::new($true))
            try {
                $brokenWriter.WriteLine('"CheminDossier";"NomFileShare";"DateDetection"')
                $dateNow = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
                foreach ($brokenPath in $allBrokenPaths) {
                    $brokenWriter.WriteLine('"' + ($brokenPath -replace '"', '""') + '";"[multi]";"' + $dateNow + '"')
                }
                $brokenWriter.Flush()
            } finally {
                $brokenWriter.Close()
            }
            Write-Log "CSV héritage cassé : $brokenCsvPath" "SUCCESS"
        }

        if ($globalStats.N1Reussis -eq 0) {
            Write-Log "Aucun scan N1 réussi. Vérifiez les droits d'accès et le mapping." "WARN"
            $WarningCount++
        }
        if ($ErrorCount -gt 0 -and $globalStats.N1Reussis -gt 0) {
            $Status = "PartialSuccess"
        } elseif ($ErrorCount -gt 0 -and $globalStats.N1Reussis -eq 0) {
            $Status = "Failed"
        }

        $mode         = if ($PSCmdlet.ParameterSetName -eq 'Mapping') { 'Mapping' } else { 'Single' }
        $psVersion    = $PSVersionTable.PSVersion.ToString()
        $parallelMode = ($PSVersionTable.PSVersion.Major -ge 7 -and $ThrottleLimit -gt 1)
        $durationSec  = [int][math]::Round(((Get-Date) - $StartTime).TotalSeconds)

        $metadata = [ordered]@{
            session_id             = $timestamp
            script_version         = "v2.2"
            vm_source              = $env:COMPUTERNAME
            compte_execution       = Get-ExecutionAccount
            mode                   = $mode
            throttle_limit         = $ThrottleLimit
            include_inherited_only = [bool]$IncludeInheritedOnly
            ps_version             = $psVersion
            parallel_mode          = $parallelMode
            mapping_csv            = if ($mode -eq 'Mapping') { $MappingCsv } else { "" }
            n1_total               = $cheminsAScanner.Count
            n1_reussis             = $globalStats.N1Reussis
            n1_refuses             = $globalStats.N1Refuses
            n1_erreurs             = $globalStats.N1Erreurs
            totaux                 = [ordered]@{
                dossiers_enumeres            = $globalStats.DossiersEnumeres
                dossiers_scannes             = ($globalStats.DossiersEnumeres - $globalStats.ErreursACL)
                dossiers_skipped_inherited   = $globalStats.DossiersSkipped
                dossiers_avec_ace_explicites = ($globalStats.DossiersEnumeres - $globalStats.DossiersSkipped - $globalStats.ErreursACL)
                dossiers_inheritance_broken  = $globalStats.DossiersInhBroken
                ace_exportees                = $globalStats.AceExportees
                duration_seconds             = $durationSec
            }
            csv_consolide          = [System.IO.Path]::GetFileName($csvConsolidePath)
            csv_par_n1             = @($csvParN1Paths | ForEach-Object { [System.IO.Path]::GetFileName($_) })
        }

        $metadata | ConvertTo-Json -Depth 5 | Out-File -FilePath $metadataPath -Encoding UTF8

        $EndTime = Get-Date
        Write-Log "Analyse des permissions terminée." "SUCCESS"
        Write-Log "N1 total : $($cheminsAScanner.Count) | réussis : $($globalStats.N1Reussis) | refusés : $($globalStats.N1Refuses) | erreurs : $($globalStats.N1Erreurs)" "INFO"
        Write-Log "Dossiers énumérés : $($globalStats.DossiersEnumeres)" "INFO"
        Write-Log "Dossiers skippés (héritage pur) : $($globalStats.DossiersSkipped)" "INFO"
        Write-Log "Dossiers à héritage cassé : $($globalStats.DossiersInhBroken)" "INFO"
        Write-Log "ACE exportées : $($globalStats.AceExportees)" "INFO"
        Write-Log "Durée : $durationSec secondes" "INFO"

        $executionChemin = if ($PSCmdlet.ParameterSetName -eq 'Mapping') { "MappingCsv:$MappingCsv" } else { $CheminUNC }
        Write-ExecutionLog -LogPath $LogPath -ScriptName $ScriptName -StartTime $StartTime -EndTime $EndTime `
            -CheminUNC $executionChemin -OutputCsv $csvConsolidePath `
            -ElementsAnalyses $ItemsScanned -ResultatsTrouves $globalStats.AceExportees `
            -WarningCount $WarningCount -ErrorCount $ErrorCount -Status $Status

        Write-Log "Fichier CSV consolidé : $csvConsolidePath" "SUCCESS"
        Write-Log "Fichier metadata      : $metadataPath" "SUCCESS"
        Write-Log "Fichier LOG           : $LogPath" "SUCCESS"
        if ($allBrokenPaths.Count -gt 0) {
            Write-Log "Fichier héritage cassé : $brokenCsvPath" "SUCCESS"
        }
        return $csvConsolidePath

    } catch {
        $Status = "Failed"
        Write-Log "Erreur finale lors de l'export consolidé : $($_.Exception.Message)" "ERROR"
        throw
    }
}