<#
.SYNOPSIS
    Produit l'inventaire d'un FileShare (mode mono-chemin ou multi-chemins via mapping CSV).

.DESCRIPTION
    Ce script analyse récursivement un ou plusieurs chemins source (local ou UNC) et produit
    un inventaire global du FileShare.

    Deux modes d'utilisation :
    - Mode Single : -CheminUNC <chemin> (rétrocompatible avec l'ancien comportement)
    - Mode Mapping : -MappingCsv <chemin du fichier FileShareMapping.csv>

    En mode Mapping, le script itère sur tous les chemins N1 listés dans le CSV
    et est résilient aux refus d'accès (continue on error).

.PARAMETER CheminUNC
    Chemin local ou UNC à analyser (mode Single).
    Exemple : \\serveur\partage

.PARAMETER MappingCsv
    Chemin vers le fichier FileShareMapping.csv listant les N1 à scanner (mode Mapping).
    Le CSV doit contenir au minimum les colonnes : CheminUNC;NomFileShare

.PARAMETER OutputPath
    Dossier de sortie (par défaut : .\Output).

.PARAMETER IncludeFileDetail
    Génère un CSV détaillé fichier par fichier avec métadonnées complètes.
    Voir documentation pour la liste des colonnes.

.PARAMETER FileDetailMinSizeMB
    Filtre les fichiers en dessous de cette taille (MB) dans le CSV de détail (défaut : 0 = tout inclure).

.EXAMPLE
    # Mode mono-chemin (rétrocompatible)
    .\Get-FileShareInventory.ps1 -CheminUNC "\\<SERVEUR_FILESHARE>\Test_Migration"

.EXAMPLE
    # Mode multi-chemins via mapping CSV
    .\Get-FileShareInventory.ps1 -MappingCsv ".\Config\FileShareMapping.Carambar.csv"

.EXAMPLE
    # Mode multi-chemins avec exécution sous compte de service (runas /netonly)
    runas /netonly /user:<DOMAINE_AD>\<COMPTE_SERVICE_LECTURE> powershell.exe
    # Dans la nouvelle fenêtre PowerShell :
    .\Get-FileShareInventory.ps1 -MappingCsv ".\Config\FileShareMapping.Carambar.csv"

.EXAMPLE
    # Mode complet avec détail fichier par fichier
    .\Get-FileShareInventory.ps1 -MappingCsv ".\Config\FileShareMappingByOU.csv" -IncludeFileDetail -FileDetailMinSizeMB 1

.NOTES
    Projet  : Carambar - Migration FileShare vers M365
    Phase   : 01 - Assessment

    Authentification :
    Pour utiliser un compte de service différent de la session courante (ex: <COMPTE_SERVICE_LECTURE>),
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

    [Parameter(Mandatory = $false, HelpMessage = "Dossier de sortie")]
    [ValidateNotNullOrEmpty()]
    [string]$OutputPath = (Join-Path -Path $PSScriptRoot -ChildPath "Output"),

    [Parameter(Mandatory = $false)]
    [PSCustomObject]$Run = $null,

    [Parameter(Mandatory = $false, HelpMessage = "Génère un CSV détaillé fichier par fichier avec métadonnées complètes")]
    [switch]$IncludeFileDetail,

    [Parameter(Mandatory = $false, HelpMessage = "Filtre les fichiers en dessous de cette taille (MB) dans le CSV de détail")]
    [ValidateRange(0, [int]::MaxValue)]
    [int]$FileDetailMinSizeMB = 0,

    [Parameter(Mandatory = $false)]
    [System.Management.Automation.PSCredential]
    [System.Management.Automation.Credential()]
    $Credential,

    [Parameter(Mandatory = $false)]
    [string]$Server = '<SERVEUR_FILESHARE>'
)

begin {
    Import-Module (Join-Path -Path $PSScriptRoot -ChildPath "Modules\Carambar.Assessment.psm1") -Force
    $outputModulePath = Join-Path -Path $PSScriptRoot -ChildPath "Modules\Carambar.Output.psm1"
    if (Test-Path $outputModulePath) { Import-Module $outputModulePath -Force }
    $smbModulePath = Join-Path -Path $PSScriptRoot -ChildPath "Modules\Carambar.SmbCredential.psm1"
    if (Test-Path $smbModulePath) {
        Import-Module $smbModulePath -Force -ErrorAction $(if ($Credential) { 'Stop' } else { 'SilentlyContinue' })
    } elseif ($Credential) {
        throw "Module Carambar.SmbCredential introuvable. Chemin attendu : $smbModulePath"
    }
    $script:smbConnected = $false

    $ScriptName = Split-Path -Leaf $PSCommandPath
    $ScriptBaseName = [System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)
    $StartTime = Get-Date
    $WarningCount = 0
    $ErrorCount = 0
    $ItemsScanned = 0
    $Status = "Success"

    function Get-HumanSize {
        param([double]$Bytes)

        if ($Bytes -ge 1TB) { return "{0:N2} TB" -f ($Bytes / 1TB) }
        elseif ($Bytes -ge 1GB) { return "{0:N2} GB" -f ($Bytes / 1GB) }
        elseif ($Bytes -ge 1MB) { return "{0:N2} MB" -f ($Bytes / 1MB) }
        elseif ($Bytes -ge 1KB) { return "{0:N2} KB" -f ($Bytes / 1KB) }
        else { return "{0:N0} octets" -f $Bytes }
    }

    function Format-CsvField {
        param($Value)
        if ($null -eq $Value) { return '""' }
        '"' + ([string]$Value).Replace('"', '""') + '"'
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
                return "$($env:USERDOMAIN)\\$($env:USERNAME)"
            }

            if (-not [string]::IsNullOrWhiteSpace($env:USERNAME)) {
                return $env:USERNAME
            }

            return "Inconnu"
        }
    }

    function Invoke-N1Scan {
        param(
            [Parameter(Mandatory = $true)][string]$CheminN1,
            [Parameter(Mandatory = $true)][string]$NomFS,
            [Parameter(Mandatory = $true)][string]$Mode,
            [Parameter(Mandatory = $true)][string]$Timestamp,
            [Parameter(Mandatory = $true)][string]$OutputPath,
            [Parameter(Mandatory = $true)][bool]$IncludeFileDetail,
            [Parameter(Mandatory = $true)][long]$FileDetailMinSizeBytes
        )

        $results = New-Object 'System.Collections.Generic.List[object]'
        $globalFileCount = 0
        $globalDirectoryCount = 0
        $globalSizeBytes = [long]0
        $hiddenSystemFileCount = 0
        $hiddenSystemSizeBytes = [long]0
        $itemsScannedLocal = 0
        $scanErrorsList = New-Object 'System.Collections.Generic.List[object]'
        $seuilLVT = 5000
        $dossierEnfantsDirects = @{}
        $detailWriter = $null
        $detailFileCount = 0
        $detailCsvPath = $null
        $analysisDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

        $safeNomFS = Get-SafeFileName -Value $NomFS

        if ($IncludeFileDetail) {
            if ($Mode -eq 'Mapping') {
                $detailCsvPath = Join-Path -Path $OutputPath -ChildPath "Inventaire_FileShare_Detail_${safeNomFS}_${Timestamp}.csv"
            }
            else {
                $detailCsvPath = Join-Path -Path $OutputPath -ChildPath "Inventaire_FileShare_Detail_${Timestamp}.csv"
            }

            try {
                $utf8BomEncoding = New-Object System.Text.UTF8Encoding($true)
                $detailWriter = New-Object System.IO.StreamWriter($detailCsvPath, $false, $utf8BomEncoding)
                $detailWriter.WriteLine('"CheminComplet";"CheminRelatif";"DossierNiveau1";"DossierParent";"NomFichier";"Extension";"TailleOctets";"TailleLisible";"DateCreation";"DateModification";"DateDernierAcces";"Attributs";"EstHidden";"EstSystem";"EstReadOnly";"EstArchive";"EstReparsePoint";"LongueurChemin";"DateAnalyse"')
            }
            catch {
                throw "Impossible de créer le CSV détaillé pour '$NomFS' : $($_.Exception.Message)"
            }
        }

        # Un seul appel non-récursif pour les dossiers de niveau 1 (rapide)
        $level1Directories = Get-ChildItem -Path $CheminN1 -Force -Directory -ErrorAction SilentlyContinue

        # Initialiser les stats par dossier de niveau 1
        $level1Stats = @{}
        $normalizedRoot = $CheminN1.TrimEnd([char]'\')
        foreach ($dir in $level1Directories) {
            $level1Stats[$dir.FullName] = @{
                Name      = $dir.Name
                FileCount = 0
                DirCount  = 0
                SizeBytes = [long]0
            }
        }

        try {
            # UN SEUL PASS récursif en streaming — énumération résiliente
            Invoke-SafeRecursiveScan -RootPath $CheminN1 -ErrorCollection $scanErrorsList `
                -NomFileShare $NomFS -ProgressActivity "Inventaire FileShare" | ForEach-Object {
                $itemsScannedLocal++
                $item = $_

                # Wrapper obligatoire : tout accès aux propriétés en try/catch individuel
                try {
                    $relativePath = $item.FullName
                    if ($item.FullName.StartsWith($normalizedRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
                        $relativePath = $item.FullName.Substring($normalizedRoot.Length).TrimStart([char]'\')
                    }

                    $level1Name = ($relativePath -split '\\')[0]
                    $level1Key = Join-Path $normalizedRoot $level1Name

                    # Détection attributs Hidden / System
                    $estHiddenOuSystem = ($item.Attributes -band [System.IO.FileAttributes]::Hidden) -or ($item.Attributes -band [System.IO.FileAttributes]::System)

                    # Comptage des enfants directs du dossier parent (pour LVT)
                    $parentDir = Split-Path -Parent $item.FullName
                    if ($parentDir) {
                        if (-not $dossierEnfantsDirects.ContainsKey($parentDir)) {
                            $dossierEnfantsDirects[$parentDir] = 0
                        }
                        $dossierEnfantsDirects[$parentDir]++
                    }

                    $isDir = ($item -is [System.IO.DirectoryInfo])

                    if ($isDir) {
                        $globalDirectoryCount++
                        if ($level1Stats.ContainsKey($level1Key)) {
                            $level1Stats[$level1Key].DirCount++
                        }
                    }
                    else {
                        $globalFileCount++
                        $globalSizeBytes += $item.Length
                        if ($estHiddenOuSystem) {
                            $hiddenSystemFileCount++
                            $hiddenSystemSizeBytes += $item.Length
                        }
                        if ($level1Stats.ContainsKey($level1Key)) {
                            $level1Stats[$level1Key].FileCount++
                            $level1Stats[$level1Key].SizeBytes += $item.Length
                        }

                        if ($detailWriter -and ($FileDetailMinSizeBytes -le 0 -or $item.Length -ge $FileDetailMinSizeBytes)) {
                            $attributes = $item.Attributes
                            $fileExtension = [string]$item.Extension
                            $detailLine = (@(
                                    Format-CsvField $item.FullName
                                    Format-CsvField $relativePath
                                    Format-CsvField $level1Name
                                    Format-CsvField (Split-Path -Parent $item.FullName)
                                    Format-CsvField $item.Name
                                    Format-CsvField $fileExtension.ToLowerInvariant()
                                    Format-CsvField $item.Length
                                    Format-CsvField (Get-HumanSize -Bytes $item.Length)
                                    Format-CsvField ($item.CreationTime.ToString("yyyy-MM-dd HH:mm:ss"))
                                    Format-CsvField ($item.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss"))
									Format-CsvField ($item.LastAccessTime.ToString("yyyy-MM-dd HH:mm:ss"))
                                    Format-CsvField $attributes.ToString()
                                    Format-CsvField ([bool]($attributes -band [System.IO.FileAttributes]::Hidden))
                                    Format-CsvField ([bool]($attributes -band [System.IO.FileAttributes]::System))
                                    Format-CsvField ([bool]($attributes -band [System.IO.FileAttributes]::ReadOnly))
                                    Format-CsvField ([bool]($attributes -band [System.IO.FileAttributes]::Archive))
                                    Format-CsvField ([bool]($attributes -band [System.IO.FileAttributes]::ReparsePoint))
                                    Format-CsvField $item.FullName.Length
                                    Format-CsvField $analysisDate
                                )) -join ';'
                            $detailWriter.WriteLine($detailLine)
                            $detailFileCount++
                        }
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
        }
        finally {
            if ($detailWriter) {
                $detailWriter.Flush()
                $detailWriter.Close()
                $detailWriter.Dispose()
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

        $results.Add([PSCustomObject]@{
            NomFileShare    = $NomFS
            TypeLigne       = "Global"
            DossierNiveau1  = "[TOTAL]"
            CheminAnalyse   = $CheminN1
            NombreFichiers  = $globalFileCount
            NombreDossiers  = $globalDirectoryCount
            TailleOctets    = $globalSizeBytes
            TailleLisible   = Get-HumanSize -Bytes $globalSizeBytes
            DateAnalyse     = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            ResultatAnalyse = "Inventaire global"
        })

        # Construire les résultats niveau 1 à partir des données déjà collectées lors du premier pass
        # (aucun second parcours récursif nécessaire — critique sur 40 TO de données)
        foreach ($key in $level1Stats.Keys) {
            $stat = $level1Stats[$key]
            $results.Add([PSCustomObject]@{
                NomFileShare    = $NomFS
                TypeLigne       = "DossierNiveau1"
                DossierNiveau1  = $stat.Name
                CheminAnalyse   = $key
                NombreFichiers  = $stat.FileCount
                NombreDossiers  = $stat.DirCount
                TailleOctets    = $stat.SizeBytes
                TailleLisible   = Get-HumanSize -Bytes $stat.SizeBytes
                DateAnalyse     = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                ResultatAnalyse = "Inventaire dossier niveau 1"
            })
        }

        if ($level1Directories.Count -eq 0 -and $globalFileCount -gt 0) {
            Write-Log "Aucun sous-dossier de premier niveau détecté pour $NomFS. Le partage semble contenir des fichiers directement à la racine." "WARN"
            $script:WarningCount++
        }

        return @{
            Results            = $results
            FileCount          = $globalFileCount
            DirCount           = $globalDirectoryCount
            SizeBytes          = $globalSizeBytes
            HiddenSystemCount  = $hiddenSystemFileCount
            HiddenSystemBytes  = $hiddenSystemSizeBytes
            LVTFolders         = $dossierEnfantsDirects
            DetailCsvPath      = $detailCsvPath
            DetailFileCount    = $detailFileCount
            ItemsScanned       = $itemsScannedLocal
        }
    }

    $script:RunErrorsPath = $null
    $script:RunMode       = $false
    Write-Log "Initialisation de l'inventaire du FileShare"
    Initialize-OutputPath -Path $OutputPath

    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $csvConsolidePath = Join-Path -Path $OutputPath -ChildPath "Inventaire_FileShare_${timestamp}.csv"
    $metadataPath = Join-Path -Path $OutputPath -ChildPath "Inventaire_FileShare_Metadata_${timestamp}.json"
    $LogPath = Join-Path -Path $OutputPath -ChildPath "$ScriptBaseName`_$timestamp.log"
    Set-LogFile -Path $LogPath
    # Routing: si $Run fourni → chemins hiérarchiques ; sinon → mode legacy
    if ($null -ne $Run) {
        $script:RunErrorsPath = $Run.Errors
        $script:RunMode       = $true
        $OutputPath           = $Run.Csv
        $metadataPath         = Join-Path $Run.Metadata "Inventaire_FileShare.json"
        $LogPath              = Join-Path $Run.Logs "Get-FileShareInventory.log"
        $csvConsolidePath     = Join-Path $Run.Csv "Inventaire_FileShare.csv"
        Set-LogFile -Path $LogPath
    }
    Write-Log "=== Démarrage $ScriptName ===" "INFO"
    Write-Log "VM source : $env:COMPUTERNAME" "INFO"
    Write-Log "Compte d'exécution : $(Get-ExecutionAccount)" "INFO"
    Write-Log "OutputPath : $OutputPath" "INFO"
    $fileDetailMinSizeBytes = [long]($FileDetailMinSizeMB * 1MB)

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
        TotalFichiers = [long]0
        TotalDossiers = [long]0
        TotalOctets   = [long]0
        N1Reussis     = 0
        N1Refuses     = 0
        N1Erreurs     = 0
    }
}

process {
    if ($Credential) {
        $smbResult = Connect-CarambarFileShare -Server $Server -Credential $Credential
        Test-FileShareIdentity -Server $Server -ExpectedUserName $Credential.UserName
        $script:smbConnected = -not $smbResult.AlreadyConnected
    }

    Write-Log "Mode d'exécution : $($PSCmdlet.ParameterSetName)"

    foreach ($entry in $cheminsAScanner) {
        $cheminN1 = $entry.CheminUNC
        $nomFS = $entry.NomFileShare
        $safeNomFS = Get-SafeFileName -Value $nomFS

        Write-Log "=== Démarrage scan : $nomFS ($cheminN1) ===" "INFO"

        # Vérification accès AVANT le scan (évite de partir sur un dossier refusé)
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

        # Scan réel du N1
        try {
            $statsN1 = Invoke-N1Scan -CheminN1 $cheminN1 -NomFS $nomFS -Mode $PSCmdlet.ParameterSetName -Timestamp $timestamp -OutputPath $OutputPath -IncludeFileDetail ([bool]$IncludeFileDetail) -FileDetailMinSizeBytes $fileDetailMinSizeBytes

            # Export CSV par N1 (un fichier par dossier scanné)
            if ($PSCmdlet.ParameterSetName -eq 'Mapping') {
                $csvN1Name = "Inventaire_FileShare_${safeNomFS}_${timestamp}.csv"
            }
            else {
                $csvN1Name = "Inventaire_FileShare_${timestamp}.csv"
            }

            $csvN1Path = Join-Path -Path $OutputPath -ChildPath $csvN1Name
            $statsN1.Results | Export-Csv -Path $csvN1Path -NoTypeInformation -Delimiter ';' -Encoding UTF8
            Write-Log "CSV par N1 généré : $csvN1Path" "SUCCESS"
            $csvParN1.Add($csvN1Name)

            if ($statsN1.DetailCsvPath) {
                Write-Log "Nombre de fichiers exportés dans le détail ($nomFS) : $($statsN1.DetailFileCount)" "INFO"
                Write-Log "CSV détaillé généré : $($statsN1.DetailCsvPath)" "SUCCESS"
            }

            # Agrégation pour le CSV consolidé
            $allResults.AddRange($statsN1.Results)

            # Stats globales
            $globalStats.TotalFichiers += [long]$statsN1.FileCount
            $globalStats.TotalDossiers += [long]$statsN1.DirCount
            $globalStats.TotalOctets += [long]$statsN1.SizeBytes
            $globalStats.N1Reussis++
            $ItemsScanned += [int]$statsN1.ItemsScanned

            # Signalement des fichiers Hidden/System
            if ($statsN1.HiddenSystemCount -gt 0) {
                Write-Log "Fichiers Hidden/System détectés sur $nomFS : $($statsN1.HiddenSystemCount) ($(Get-HumanSize -Bytes $statsN1.HiddenSystemBytes))" "WARN"
                Write-Log "ATTENTION : ShareGate ignore les fichiers et dossiers Hidden/System lors de la migration." "WARN"
                $WarningCount += 2
            }

            # Détection des dossiers dépassant le List View Threshold (> 5000 éléments directs)
            $seuilLVT = 5000
            $dossiersLVT = $statsN1.LVTFolders.GetEnumerator() | Where-Object { $_.Value -gt $seuilLVT }
            $nbDossiersLVT = ($dossiersLVT | Measure-Object).Count
            if ($nbDossiersLVT -gt 0) {
                Write-Log "Dossiers dépassant le List View Threshold ($seuilLVT items directs) sur $nomFS : $nbDossiersLVT" "WARN"
                $WarningCount++
                $CsvLVTPath = Join-Path -Path $OutputPath -ChildPath "Dossiers_LVT_${safeNomFS}_$timestamp.csv"
                $dossiersLVT | Sort-Object Value -Descending | ForEach-Object {
                    [PSCustomObject]@{
                        NomFileShare         = $nomFS
                        CheminDossier        = $_.Key
                        NombreElementsDirects = $_.Value
                        Seuil                = $seuilLVT
                        Depassement          = $_.Value - $seuilLVT
                        Severite             = "WARN"
                        DateAnalyse          = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                        ResultatAnalyse      = "Dossier dépassant le List View Threshold SharePoint Online"
                    }
                } | Export-Csv -Path $CsvLVTPath -NoTypeInformation -Delimiter ';' -Encoding UTF8
                Write-Log "Fichier CSV LVT : $CsvLVTPath" "INFO"
            }
        }
        catch [System.UnauthorizedAccessException] {
            Write-Log "Access denied en cours de scan sur $nomFS - skip partiel" "WARN"
            $WarningCount++
            $globalStats.N1Refuses++
            continue
        }
        catch {
            Write-Log "Erreur scan $nomFS : $($_.Exception.Message)" "ERROR"
            $ErrorCount++
            $globalStats.N1Erreurs++
            continue
        }
    }
}

end {
    if ($script:smbConnected) {
        try { Disconnect-CarambarFileShare -Server $Server } catch { Write-Verbose "Disconnect-CarambarFileShare: $($_.Exception.Message)" }
        $script:smbConnected = $false
    }
    try {
        # Ligne TOTAL globale ajoutée au consolidé (pour Export-AssessmentReport.ps1)
        $ligneTotal = [PSCustomObject]@{
            NomFileShare    = "[CONSOLIDE]"
            TypeLigne       = "Global"
            DossierNiveau1  = "[TOTAL]"
            CheminAnalyse   = "Multiple sources"
            NombreFichiers  = $globalStats.TotalFichiers
            NombreDossiers  = $globalStats.TotalDossiers
            TailleOctets    = $globalStats.TotalOctets
            TailleLisible   = Get-HumanSize -Bytes $globalStats.TotalOctets
            DateAnalyse     = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            ResultatAnalyse = "Inventaire global consolidé"
        }

        # CSV consolidé (toutes les lignes de tous les N1)
        @($ligneTotal) + $allResults | Export-Csv -Path $csvConsolidePath -NoTypeInformation -Delimiter ';' -Encoding UTF8

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
            duration_seconds = $durationSeconds
            totaux           = [ordered]@{
                fichiers      = $globalStats.TotalFichiers
                dossiers      = $globalStats.TotalDossiers
                octets        = $globalStats.TotalOctets
                taille_lisible = Get-HumanSize -Bytes $globalStats.TotalOctets
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

        Write-Log "Inventaire consolidé terminé." "SUCCESS"
        Write-Log "N1 total : $($cheminsAScanner.Count) | réussis : $($globalStats.N1Reussis) | refusés : $($globalStats.N1Refuses) | erreurs : $($globalStats.N1Erreurs)" "INFO"
        Write-Log "Nombre total de fichiers : $($globalStats.TotalFichiers)" "INFO"
        Write-Log "Nombre total de dossiers : $($globalStats.TotalDossiers)" "INFO"
        Write-Log "Volume total : $(Get-HumanSize -Bytes $globalStats.TotalOctets)" "INFO"

        $EndTime = Get-Date
        Write-Progress -Activity "Inventaire du FileShare" -Completed

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
