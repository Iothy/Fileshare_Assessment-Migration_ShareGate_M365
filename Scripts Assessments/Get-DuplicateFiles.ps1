<#
.SYNOPSIS
    Détecte les fichiers dupliqués dans un FileShare (mode mono-chemin ou multi-chemins via mapping CSV).

.DESCRIPTION
    Ce script analyse récursivement un ou plusieurs chemins source (local ou UNC), calcule les empreintes
    des fichiers et identifie les doublons.

    Deux modes d'utilisation :
    - Mode Single : -CheminUNC <chemin> (rétrocompatible)
    - Mode Mapping : -MappingCsv <chemin du fichier FileShareMapping.csv>

    En mode Mapping, chaque N1 est traité indépendamment pour détecter les doublons intra-N1.

.PARAMETER CheminUNC
    Chemin local ou UNC à analyser (mode Single).
    Exemple : \\serveur\partage

.PARAMETER MappingCsv
    Chemin vers le fichier FileShareMapping.csv listant les N1 à scanner (mode Mapping).
    Le CSV doit contenir au minimum les colonnes : CheminUNC;NomFileShare

.PARAMETER Algorithm
    Algorithme de hash utilisé.
    Valeurs possibles : MD5, SHA1, SHA256, SHA384, SHA512
    Valeur par défaut : SHA256

.PARAMETER MinSizeMB
    Taille minimale en MB pour considérer un fichier comme candidat doublon.
    Valeur par défaut : 50 MB (sweet spot ROI/perf identifié sur le périmètre DSI PrimaGAZ).

    Recommandations selon contexte :
    - 1 MB   : audit exhaustif (très long, beaucoup de bruit, à utiliser pour conformité uniquement)
    - 10 MB  : audit étendu (1-2h sur ~600 GB)
    - 50 MB  : sweet spot livrable client (20-45 min sur ~600 GB) - DÉFAUT
    - 100 MB : analyse rapide top doublons (5-10 min)
    - 500 MB : analyse focus archives/backups (1-2 min)

.PARAMETER OutputPath
    Dossier de sortie.
    Par défaut : .\Output (dans le dossier du script)

.EXAMPLE
    # Mode mono-chemin (rétrocompatible) — seuil par défaut 50 MB
    .\Get-DuplicateFiles.ps1 -CheminUNC \\serveur\partage

.EXAMPLE
    # Mode mono-chemin avec seuil 50 MB (sweet spot livrable client sur ~600 GB, 20-45 min)
    .\Get-DuplicateFiles.ps1 -CheminUNC \\serveur\partage -MinSizeMB 50

.EXAMPLE
    # Mode mono-chemin avec filtre taille exhaustif (attention : très long sur grands volumes)
    .\Get-DuplicateFiles.ps1 -CheminUNC \\serveur\partage -MinSizeMB 1

.EXAMPLE
    # Mode multi-chemins via mapping CSV
    .\Get-DuplicateFiles.ps1 -MappingCsv ".\Config\FileShareMapping.PrimaGAZ.csv"

.EXAMPLE
    # Mode multi-chemins avec exécution sous compte de service (runas /netonly)
    runas /netonly /user:shvenergy.corp\FR_PG_SVC_FileShare_Read powershell.exe
    # Dans la nouvelle fenêtre PowerShell :
    .\Get-DuplicateFiles.ps1 -MappingCsv ".\Config\FileShareMapping.PrimaGAZ.csv"

.NOTES
    Projet  : PrimaGAZ - Migration FileShare vers M365
    Phase   : 01 - Assessment

    Authentification :
    Pour utiliser un compte de service différent de la session courante (ex: FR_PG_SVC_FileShare_Read),
    lancer une session PowerShell via : runas /netonly /user:DOMAIN\user powershell.exe
    Les credentials seront utilisés uniquement pour les accès SMB.
#>

[CmdletBinding(DefaultParameterSetName = 'Single')]
param(
    [Parameter(Mandatory, ParameterSetName = 'Single', HelpMessage = "Chemin local ou UNC à analyser")]
    [ValidateNotNullOrEmpty()]
    [string]$CheminUNC,

    [Parameter(Mandatory, ParameterSetName = 'Mapping', HelpMessage = "Chemin du fichier FileShareMapping.csv (colonnes : CheminUNC;NomFileShare;...)")]
    [ValidateNotNullOrEmpty()]
    [string]$MappingCsv,

    [Parameter(Mandatory = $false, HelpMessage = "Algorithme de hash à utiliser")]
    [ValidateSet("MD5", "SHA1", "SHA256", "SHA384", "SHA512")]
    [string]$Algorithm = "SHA256",

    [Parameter(Mandatory = $false, HelpMessage = "Taille minimale en MB pour considérer un fichier comme candidat doublon")]
    [ValidateRange(0, 1048576)]
    [int]$MinSizeMB = 100,

    [Parameter(Mandatory = $false, HelpMessage = "Dossier de sortie")]
    [ValidateNotNullOrEmpty()]
    [string]$OutputPath = (Join-Path -Path $PSScriptRoot -ChildPath "Output"),

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

    $ScriptName = Split-Path -Leaf $PSCommandPath
    $ScriptBaseName = [System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)
    $StartTime = Get-Date
    $WarningCount = 0
    $ErrorCount = 0
    $ItemsScanned = 0
    $Status = "Success"

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

        $trimmed = $Path.TrimEnd([char]'\', [char]'/')
        if ([string]::IsNullOrWhiteSpace($trimmed)) { return "FileShare" }

        $parts = $trimmed -split '[/\\]'
        $nonEmpty = @($parts | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        if ($nonEmpty.Count -eq 0) { return "FileShare" }

        return $nonEmpty[-1]
    }

    function Get-ExecutionAccount {
        try {
            if ($PSVersionTable.PSEdition -eq 'Desktop') {
                return [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
            }
        }
        catch {
            # fallback ci-dessous
        }

        try {
            return (whoami)
        }
        catch {
            if (-not [string]::IsNullOrWhiteSpace($env:USERDOMAIN) -and -not [string]::IsNullOrWhiteSpace($env:USERNAME)) {
                return "$($env:USERDOMAIN)\$($env:USERNAME)"
            }

            if (-not [string]::IsNullOrWhiteSpace($env:USERNAME)) {
                return $env:USERNAME
            }

            return "Inconnu"
        }
    }

    function Export-NoResult {
        param(
            [string]$Path,
            [string]$NomFileShare
        )

        @(
            [PSCustomObject]@{
                GroupeDoublonId      = ""
                Hash                 = ""
                NombreOccurrences    = ""
                TailleOctets         = ""
                TailleMB             = ""
                NomFichier           = ""
                CheminComplet        = ""
                DerniereModification = ""
                DateAnalyse          = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                ResultatAnalyse      = "Aucun doublon détecté"
                NomFileShare         = $NomFileShare
            }
        ) | Export-Csv -Path $Path -NoTypeInformation -Delimiter ";" -Encoding UTF8
    }

    function Invoke-N1DuplicatesScan {
        param(
            [Parameter(Mandatory = $true)][string]$CheminN1,
            [Parameter(Mandatory = $true)][string]$NomFS,
            [Parameter(Mandatory = $true)][string]$Algorithm,
            [Parameter(Mandatory = $true)][int]$MinSizeMB
        )

        $minBytes = $MinSizeMB * 1MB
        $results = New-Object 'System.Collections.Generic.List[object]'
        $sizeGroups = @{}
        $groupId = 0
        $hashErrorCount = 0
        $itemsScannedLocal = 0
        $scanErrorsList = New-Object 'System.Collections.Generic.List[object]'
        $safeNomFS = Get-SafeFileName -Value $NomFS

        # Pass 1 : regroupement rapide par taille (énumération résiliente)
        Invoke-SafeRecursiveScan -RootPath $CheminN1 -FilesOnly -ErrorCollection $scanErrorsList `
            -NomFileShare $NomFS -ProgressActivity "Analyse des doublons (Pass 1 - taille)" |
            ForEach-Object {
                $item = $_
                try {
                    if ($item.Length -ge $minBytes) {
                        $itemsScannedLocal++
                        if (-not $sizeGroups.ContainsKey($item.Length)) {
                            $sizeGroups[$item.Length] = New-Object 'System.Collections.Generic.List[object]'
                        }
                        $sizeGroups[$item.Length].Add($item)
                    }
                }
                catch {
                    $errEntry = [PSCustomObject]@{
                        NomFileShare  = $NomFS
                        Chemin        = if ($item) { $item.FullName } else { '<unknown>' }
                        TypeErreur    = 'ProcessingError'
                        ExceptionType = $_.Exception.GetType().Name
                        MessageErreur = $_.Exception.Message
                        DateDetection = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                    }
                    $scanErrorsList.Add($errEntry)
                    Write-Log "Erreur traitement objet : $($item.FullName) [$($_.Exception.GetType().Name)] $($_.Exception.Message)" "WARN"
                }
            }

        # Pass 2 : hash sur les seuls groupes de taille avec >=2 candidats
        $sizeKeys = @($sizeGroups.Keys)
        $candidateGroupsTotal = $sizeKeys.Count
        $groupIndex = 0

        foreach ($size in $sizeKeys) {
            $groupIndex++
            if ($groupIndex % 100 -eq 0 -or $groupIndex -eq $candidateGroupsTotal) {
                Write-Progress -Activity "Analyse des doublons" -Status "Groupe taille $groupIndex / $candidateGroupsTotal" -PercentComplete ([math]::Min(100, [math]::Round(($groupIndex / [math]::Max(1, $candidateGroupsTotal)) * 100)))
            }

            $candidates = $sizeGroups[$size]
            if ($candidates.Count -lt 2) { continue }

            $hashedEntries = New-Object 'System.Collections.Generic.List[object]'
            foreach ($file in $candidates) {
                try {
                    $hash = (Get-FileHash -Path $file.FullName -Algorithm $Algorithm -ErrorAction Stop).Hash
                    $hashedEntries.Add([PSCustomObject]@{
                        File = $file
                        Hash = $hash
                    })
                }
                catch {
                    $hashErrorCount++
                    $errEntry = [PSCustomObject]@{
                        NomFileShare  = $NomFS
                        Chemin        = $file.FullName
                        TypeErreur    = 'HashError'
                        ExceptionType = $_.Exception.GetType().Name
                        MessageErreur = $_.Exception.Message
                        DateDetection = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                    }
                    $scanErrorsList.Add($errEntry)
                    Write-Verbose "Hash KO : $($file.FullName) - $($_.Exception.Message)"
                }
            }

            $duplicateGroups = $hashedEntries | Group-Object Hash | Where-Object { $_.Count -ge 2 }

            foreach ($duplicateGroup in $duplicateGroups) {
                $groupId++
                foreach ($entry in $duplicateGroup.Group) {
                    $file = $entry.File
                    $results.Add([PSCustomObject]@{
                        GroupeDoublonId      = $groupId
                        Hash                 = $entry.Hash
                        NombreOccurrences    = $duplicateGroup.Count
                        TailleOctets         = $file.Length
                        TailleMB             = [math]::Round(($file.Length / 1MB), 2)
                        NomFichier           = $file.Name
                        CheminComplet        = $file.FullName
                        DerniereModification = $file.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss")
                        DateAnalyse          = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                        ResultatAnalyse      = "Doublon détecté"
                        NomFileShare         = $NomFS
                    })
                }
            }
        }

        # === Reporting unifié des erreurs de scan ===
        if ($scanErrorsList.Count -gt 0) {
            $denied = @($scanErrorsList | Where-Object { $_.TypeErreur -eq 'AccessDenied' })
            $others = @($scanErrorsList | Where-Object { $_.TypeErreur -ne 'AccessDenied' })
            $errBase   = if ($script:RunErrorsPath) { $script:RunErrorsPath } else { $OutputPath }
            $errSuffix = if ($script:RunMode)       { "" }                  else { "_${timestamp}" }

            if ($denied.Count -gt 0) {
                Write-Log "Sous-dossiers/objets refusés sur $NomFS : $($denied.Count)" "WARN"
                $deniedCsv = Join-Path $errBase "AccessDenied_${safeNomFS}${errSuffix}.csv"
                $denied | Export-Csv -Path $deniedCsv -NoTypeInformation -Delimiter ';' -Encoding UTF8
                Write-Log "CSV des refus : $deniedCsv" "WARN"
                $denied | Select-Object -First 5 | ForEach-Object { Write-Log "  → Refusé : $($_.Chemin) [$($_.ExceptionType)]" "WARN" }
                if ($denied.Count -gt 5) {
                    Write-Log "  → ... et $($denied.Count - 5) autre(s) refus (voir CSV)" "WARN"
                }
                $script:WarningCount += $denied.Count
            }

            if ($others.Count -gt 0) {
                Write-Log "Autres erreurs sur $NomFS : $($others.Count)" "WARN"
                $errCsv = Join-Path $errBase "ScanErrors_${safeNomFS}${errSuffix}.csv"
                $others | Export-Csv -Path $errCsv -NoTypeInformation -Delimiter ';' -Encoding UTF8
                Write-Log "CSV des erreurs : $errCsv" "WARN"
                $others | Select-Object -First 3 | ForEach-Object { Write-Log "  → $($_.ExceptionType) : $($_.Chemin)" "WARN" }
                $script:WarningCount += $others.Count
            }
        }

        return @{
            Results          = $results
            ItemsScanned     = $itemsScannedLocal
            HashErrors       = $hashErrorCount
            DuplicateGroups  = $groupId
        }
    }

    $script:RunErrorsPath = $null
    $script:RunMode       = $false
    Write-Log "Initialisation de l'analyse des doublons"
    Initialize-OutputPath -Path $OutputPath

    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $csvConsolidePath = Join-Path -Path $OutputPath -ChildPath "FichiersDupliques_${timestamp}.csv"
    $metadataPath = Join-Path -Path $OutputPath -ChildPath "FichiersDupliques_Metadata_${timestamp}.json"
    $LogPath = Join-Path -Path $OutputPath -ChildPath "$ScriptBaseName`_$timestamp.log"
    Set-LogFile -Path $LogPath
    # Routing: si $Run fourni → chemins hiérarchiques ; sinon → mode legacy
    if ($null -ne $Run) {
        $script:RunErrorsPath = $Run.Errors
        $script:RunMode       = $true
        $OutputPath           = $Run.Csv
        $metadataPath         = Join-Path $Run.Metadata "FichiersDupliques.json"
        $LogPath              = Join-Path $Run.Logs "Get-DuplicateFiles.log"
        $csvConsolidePath     = Join-Path $Run.Csv "FichiersDupliques.csv"
        Set-LogFile -Path $LogPath
    }
    Write-Log "=== Démarrage $ScriptName ===" "INFO"
    Write-Log "VM source : $env:COMPUTERNAME" "INFO"
    Write-Log "Compte d'exécution : $(Get-ExecutionAccount)" "INFO"
    Write-Log "OutputPath : $OutputPath" "INFO"

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
    }
    else {
        Test-SourcePath -Path $CheminUNC
        $cheminsAScanner.Add([PSCustomObject]@{
            CheminUNC    = $CheminUNC
            NomFileShare = Get-NomFileShareFromPath -Path $CheminUNC
        })
    }

    $allResults = New-Object 'System.Collections.Generic.List[object]'
    $csvParN1 = New-Object 'System.Collections.Generic.List[string]'

    $globalStats = @{
        N1Reussis      = 0
        N1Refuses      = 0
        N1Erreurs      = 0
        HashErrors     = 0
        GroupesDoublon = 0
    }
}

process {
    if ($Credential) {
        $smbResult = Connect-PrimaGazFileShare -Server $Server -Credential $Credential
        Test-FileShareIdentity -Server $Server -ExpectedUserName $Credential.UserName
        $script:smbConnected = -not $smbResult.AlreadyConnected
    }

    Write-Log "Mode d'exécution : $($PSCmdlet.ParameterSetName)"
    Write-Log "Algorithme utilisé : $Algorithm"
    Write-Log "Taille minimale : $MinSizeMB MB"

    foreach ($entry in $cheminsAScanner) {
        $cheminN1 = $entry.CheminUNC
        $nomFS = $entry.NomFileShare
        $safeNomFS = Get-SafeFileName -Value $nomFS

        Write-Log "=== Démarrage scan doublons : $nomFS ($cheminN1) ===" "INFO"

        # Vérification d'accès en amont
        try {
            $null = Get-Item -Path $cheminN1 -ErrorAction Stop
        }
        catch [System.UnauthorizedAccessException] {
            Write-Log "ACCESS DENIED sur $nomFS - skip (compte sans droits sur ce share)" "WARN"
            $WarningCount++
            $globalStats.N1Refuses++
            continue
        }
        catch {
            Write-Log "Chemin inaccessible : $nomFS - $($_.Exception.Message)" "ERROR"
            $ErrorCount++
            $globalStats.N1Erreurs++
            continue
        }

        try {
            $scanN1 = Invoke-N1DuplicatesScan -CheminN1 $cheminN1 -NomFS $nomFS -Algorithm $Algorithm -MinSizeMB $MinSizeMB

            $ItemsScanned += [int]$scanN1.ItemsScanned
            $globalStats.HashErrors += [int]$scanN1.HashErrors
            $globalStats.GroupesDoublon += [int]$scanN1.DuplicateGroups

            if ($PSCmdlet.ParameterSetName -eq 'Mapping') {
                $csvN1Name = "FichiersDupliques_${safeNomFS}_${timestamp}.csv"
            }
            else {
                $csvN1Name = "FichiersDupliques_${timestamp}.csv"
            }

            $csvN1Path = Join-Path -Path $OutputPath -ChildPath $csvN1Name
            if ($scanN1.Results.Count -gt 0) {
                $scanN1.Results |
                    Sort-Object GroupeDoublonId, CheminComplet |
                    Export-Csv -Path $csvN1Path -NoTypeInformation -Delimiter ";" -Encoding UTF8
            }
            else {
                Export-NoResult -Path $csvN1Path -NomFileShare $nomFS
            }

            Write-Log "CSV par N1 généré : $csvN1Path" "SUCCESS"
            $csvParN1.Add($csvN1Name)

            if ($scanN1.HashErrors -gt 0) {
                Write-Log "Erreurs hash non bloquantes sur $nomFS : $($scanN1.HashErrors)" "WARN"
                $WarningCount++
            }

            $allResults.AddRange($scanN1.Results)
            $globalStats.N1Reussis++
        }
        catch [System.UnauthorizedAccessException] {
            Write-Log "Access denied en cours de scan sur $nomFS - skip partiel" "WARN"
            $WarningCount++
            $globalStats.N1Refuses++
            continue
        }
        catch {
            Write-Log "Erreur scan doublons $nomFS : $($_.Exception.Message)" "ERROR"
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
        if ($allResults.Count -gt 0) {
            $allResults |
                Sort-Object GroupeDoublonId, CheminComplet |
                Export-Csv -Path $csvConsolidePath -NoTypeInformation -Delimiter ";" -Encoding UTF8
        }
        else {
            Export-NoResult -Path $csvConsolidePath -NomFileShare "[CONSOLIDE]"
        }

        $mode = if ($PSCmdlet.ParameterSetName -eq 'Mapping') { 'Mapping' } else { 'Single' }
        $durationSeconds = [int][math]::Round(((Get-Date) - $StartTime).TotalSeconds)
        $metadata = [ordered]@{
            session_id       = $timestamp
            vm_source        = $env:COMPUTERNAME
            compte_execution = Get-ExecutionAccount
            mode             = $mode
            mapping_csv      = if ($mode -eq 'Mapping') { $MappingCsv } else { "" }
            n1_total         = $cheminsAScanner.Count
            n1_reussis       = $globalStats.N1Reussis
            n1_refuses       = $globalStats.N1Refuses
            n1_erreurs       = $globalStats.N1Erreurs
            erreurs_hash     = $globalStats.HashErrors
            groupes_doublons = $globalStats.GroupesDoublon
            duration_seconds = $durationSeconds
            totaux           = [ordered]@{
                fichiers_dupliques = $allResults.Count
            }
            csv_consolide    = [System.IO.Path]::GetFileName($csvConsolidePath)
            csv_par_n1       = @($csvParN1)
        }

        $metadata | ConvertTo-Json -Depth 5 | Out-File -FilePath $metadataPath -Encoding UTF8

        if ($globalStats.N1Reussis -eq 0) {
            Write-Log "Aucun scan N1 réussi. Vérifiez les droits d'accès et le mapping." "WARN"
            $WarningCount++
        }

        if ($ErrorCount -gt 0 -and $globalStats.N1Reussis -gt 0) {
            $Status = "PartialSuccess"
        }
        elseif ($ErrorCount -gt 0 -and $globalStats.N1Reussis -eq 0) {
            $Status = "Failed"
        }

        $groupCount = @($allResults | Group-Object GroupeDoublonId).Count
        if ($null -eq $groupCount) { $groupCount = 0 }
        $totalDuplicateSizeMB = [math]::Round((($allResults | Measure-Object -Property TailleOctets -Sum).Sum / 1MB), 2)
        if ([double]::IsNaN($totalDuplicateSizeMB)) { $totalDuplicateSizeMB = 0 }

        if ($allResults.Count -gt 0) {
            Write-Log "Analyse terminée avec $groupCount groupe(s) de doublons détecté(s)." "WARN"; $WarningCount++
            Write-Log "Nombre total de fichiers dupliqués : $($allResults.Count)" "WARN"; $WarningCount++
            Write-Log "Volume cumulé des doublons : $totalDuplicateSizeMB MB" "WARN"; $WarningCount++
        }
        else {
            Write-Log "Aucun doublon détecté." "SUCCESS"
        }

        Write-Log "N1 total : $($cheminsAScanner.Count) | réussis : $($globalStats.N1Reussis) | refusés : $($globalStats.N1Refuses) | erreurs : $($globalStats.N1Erreurs)" "INFO"

        $EndTime = Get-Date
        Write-Progress -Activity "Analyse des doublons" -Completed

        $executionChemin = if ($PSCmdlet.ParameterSetName -eq 'Mapping') { "MappingCsv:$MappingCsv" } else { $CheminUNC }
        Write-ExecutionLog -LogPath $LogPath -ScriptName $ScriptName -StartTime $StartTime -EndTime $EndTime -CheminUNC $executionChemin -OutputCsv $csvConsolidePath -ElementsAnalyses $ItemsScanned -ResultatsTrouves $allResults.Count -WarningCount $WarningCount -ErrorCount $ErrorCount -Status $Status

        Write-Log "Fichier CSV consolidé : $csvConsolidePath" "SUCCESS"
        Write-Log "Fichier metadata : $metadataPath" "SUCCESS"
        Write-Log "Fichier LOG : $LogPath" "SUCCESS"
        return $csvConsolidePath
    }
    catch {
        $Status = "Failed"
        Write-Log "Erreur finale lors de l'export consolidé : $($_.Exception.Message)" "ERROR"
        throw
    }
}
