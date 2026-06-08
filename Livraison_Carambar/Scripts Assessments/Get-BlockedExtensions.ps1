<#
.SYNOPSIS
    Détecte les fichiers ayant des extensions de la blocklist historique SharePoint Server (mode mono-chemin ou multi-chemins via mapping CSV).

.DESCRIPTION
    Ce script analyse récursivement un ou plusieurs chemins source (FileShare local ou UNC) et identifie
    les fichiers dont l'extension figure dans la blocklist historique de SharePoint Server on-premises.

    ⚠ IMPORTANT — SharePoint Online vs SharePoint Server :
    Par défaut, SharePoint Online (SPO) ne bloque AUCUNE extension depuis 2017.
    Source MS officielle : https://support.microsoft.com/en-us/office/types-of-files-that-cannot-be-added-to-a-list-or-library-6fb5eea2-2003-47b8-9e43-3fccdbc26c30
    → « In SharePoint in Microsoft 365, there are no blocked file types »

    La liste utilisée ici correspond à la blocklist SharePoint Server on-premises (legacy).
    Les fichiers détectés sont signalés en WARN (audit) et non ERROR (bloquant) :
    il convient de vérifier la BlockDownloadFileTypePolicy du tenant SPO cible avant de
    conclure qu'un fichier ne migrera pas.

    Référence BlockDownloadFileTypePolicy : https://learn.microsoft.com/en-us/sharepoint/blocked-file-types
    Référence ShareGate : https://documentation.sharegate.com/hc/en-us/articles/115000642327

    Deux modes d'utilisation :
    - Mode Single : -CheminUNC <chemin> (rétrocompatible)
    - Mode Mapping : -MappingCsv <chemin du fichier FileShareMapping.csv>

.PARAMETER CheminUNC
    Chemin local ou UNC à analyser (mode Single).

.PARAMETER MappingCsv
    Chemin du fichier mapping CSV (mode Mapping).
    Le CSV doit contenir au minimum les colonnes : CheminUNC;NomFileShare

.PARAMETER OutputPath
    Dossier de sortie des résultats.
    Par défaut : .\Output (dans le dossier du script)

.EXAMPLE
    # Mode mono-chemin (rétrocompatible)
    .\Get-BlockedExtensions.ps1 -CheminUNC "\\serveur\partage"

.EXAMPLE
    # Mode multi-chemins (production)
    .\Get-BlockedExtensions.ps1 -MappingCsv ".\Config\FileShareMapping.Carambar.csv"

.EXAMPLE
    # Exécution avec compte de service (runas /netonly)
    runas /netonly /user:<DOMAINE_AD>\<COMPTE_SERVICE_LECTURE> powershell.exe
    # Dans la nouvelle fenêtre PowerShell :
    .\Get-BlockedExtensions.ps1 -MappingCsv ".\Config\FileShareMapping.Carambar.csv"

.NOTES
    Projet  : Carambar - Migration FileShare vers M365
    Phase   : 01 - Assessment
    Version : v1.1 — Reclassification WARN (sources officielles MS)

    Sources de référence :
    - MS Support — Types of files that cannot be added (SPO) :
      https://support.microsoft.com/en-us/office/types-of-files-that-cannot-be-added-to-a-list-or-library-6fb5eea2-2003-47b8-9e43-3fccdbc26c30
    - MS Learn — Block file types (BlockDownloadFileTypePolicy) :
      https://learn.microsoft.com/en-us/sharepoint/blocked-file-types
    - ShareGate — Files and items migration limitations :
      https://documentation.sharegate.com/hc/en-us/articles/115000642327

    La sévérité WARN (et non ERROR) reflète le fait que SPO n'a pas de liste bloquée par défaut.
    Seule une BlockDownloadFileTypePolicy configurée explicitement sur le tenant peut bloquer
    ces extensions. À vérifier avec l'administrateur M365 avant la migration.

    Authentification :
    Pour utiliser un compte de service différent de la session courante,
    lancer une session PowerShell via : runas /netonly /user:DOMAIN\user powershell.exe
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

    Write-Log "⚠ ATTENTION : la liste d'extensions utilisée correspond à SharePoint Server on-premises." "WARN"
    Write-Log "Par défaut, SharePoint Online ne bloque AUCUNE extension. Vérifier la BlockDownloadFileTypePolicy du tenant cible." "WARN"
    Write-Log "Référence : https://support.microsoft.com/en-us/office/types-of-files-that-cannot-be-added-to-a-list-or-library-6fb5eea2-2003-47b8-9e43-3fccdbc26c30" "INFO"

    # Liste des extensions issues de la blocklist SharePoint Server on-premises (legacy).
    # ⚠ SharePoint Online ne bloque aucune extension par défaut (depuis 2017).
    # Source MS : https://support.microsoft.com/en-us/office/types-of-files-that-cannot-be-added-to-a-list-or-library-6fb5eea2-2003-47b8-9e43-3fccdbc26c30
    # Source ShareGate : https://documentation.sharegate.com/hc/en-us/articles/115000642327
    $extensionsBloquees = [System.Collections.Generic.HashSet[string]]([System.StringComparer]::OrdinalIgnoreCase)
    @(
        ".ade", ".adp", ".asa", ".ashx", ".asmx", ".asp", ".bas", ".bat", ".cdx", 
        ".cer", ".chm", ".class", ".cmd", ".cnt", ".com", ".config", ".cpl", ".crt", 
        ".csh", ".der", ".exe", ".fxp", ".gadget", ".grp", ".hlp", ".hpj", ".hta", 
        ".htr", ".htw", ".ida", ".idc", ".idq", ".ins", ".isp", ".its", ".jse", 
        ".ksh", ".mad", ".maf", ".mag", ".mam", ".maq", ".mar", ".mas", ".mat", 
        ".mau", ".mav", ".maw", ".mcf", ".mda", ".mdb", ".mde", ".mdt", ".mdw", 
        ".mdz", ".msc", ".msh", ".msh1", ".msh1xml", ".msh2", ".msh2xml", ".mshxml", 
        ".msi", ".msp", ".mst", ".ops", ".ost", ".pcd", ".pif", ".pl", ".prf", 
        ".prg", ".printer", ".ps1", ".ps1xml", ".ps2", ".ps2xml", ".psc1", ".psc2", 
        ".pst", ".reg", ".rem", ".scf", ".scr", ".sct", ".shb", ".shs", ".shtm", 
        ".shtml", ".soap", ".stm", ".svc", ".vb", ".vbe", ".vbs", ".vsix", ".ws", 
        ".wsc", ".wsf", ".wsh", ".xamlx" 
		
		
    ) | ForEach-Object { [void]$extensionsBloquees.Add($_) }

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

    function New-NoResultRow {
        param(
            [string]$NomFileShare,
            [string]$Message
        )

        return [PSCustomObject]@{
            CheminComplet   = ""
            NomFichier      = ""
            Extension       = ""
            TailleOctets    = ""
            TailleMB        = ""
            Severite        = ""
            DateAnalyse     = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            ResultatAnalyse = $Message
            NomFileShare    = $NomFileShare
        }
    }

    function Invoke-N1BlockedExtensionsScan {
        param(
            [Parameter(Mandatory = $true)][string]$CheminN1,
            [Parameter(Mandatory = $true)][string]$NomFS
        )

        $results = New-Object 'System.Collections.Generic.List[object]'
        $itemsScannedLocal = 0
        $blockedCountLocal = 0
        $scanErrorsList = New-Object 'System.Collections.Generic.List[object]'
        $safeNomFS = Get-SafeFileName -Value $NomFS

        Write-Log "Début de l'analyse du chemin : $CheminN1"
        Write-Log "Nombre d'extensions bloquées référencées : $($extensionsBloquees.Count)"

        # Énumération résiliente — aucun crash possible sur objets inaccessibles
        Invoke-SafeRecursiveScan -RootPath $CheminN1 -FilesOnly -ErrorCollection $scanErrorsList `
            -NomFileShare $NomFS -ProgressActivity "Détection des extensions bloquées" | ForEach-Object {
            $itemsScannedLocal++
            $item = $_

            Write-Verbose "Analyse [$NomFS][$itemsScannedLocal] : $($item.FullName)"

            # Wrapper obligatoire : tout accès aux propriétés en try/catch individuel
            try {
                $extension = $item.Extension

                if ($extensionsBloquees.Contains($extension)) {
                    $results.Add([PSCustomObject]@{
                            CheminComplet   = $item.FullName
                            NomFichier      = $item.Name
                            Extension       = $extension
                            TailleOctets    = $item.Length
                            TailleMB        = [math]::Round(($item.Length / 1MB), 2)
                            Severite        = "WARN"
                            DateAnalyse     = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                            ResultatAnalyse = "Extension à auditer (blocklist historique SP Server) — vérifier BlockDownloadFileTypePolicy du tenant SPO avant de bloquer la migration"
                            NomFileShare    = $NomFS
                        })
                    $blockedCountLocal++
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

        Write-Log "Nombre de fichiers analysés ($NomFS) : $itemsScannedLocal"

        return @{
            Results = $results
            Stats   = @{
                ItemsScanned = $itemsScannedLocal
                BlockedCount = $blockedCountLocal
            }
        }
    }

    $script:RunErrorsPath = $null
    $script:RunMode       = $false
    Write-Log "Initialisation de la détection des extensions bloquées par SharePoint Online"
    Initialize-OutputPath -Path $OutputPath

    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $CsvPathConsolide = Join-Path -Path $OutputPath -ChildPath "ExtensionsBloquees_$timestamp.csv"
    $MetadataPath = Join-Path -Path $OutputPath -ChildPath "ExtensionsBloquees_Metadata_$timestamp.json"
    $LogPath = Join-Path -Path $OutputPath -ChildPath "$ScriptBaseName`_$timestamp.log"
    Set-LogFile -Path $LogPath
    # Routing: si $Run fourni → chemins hiérarchiques ; sinon → mode legacy
    if ($null -ne $Run) {
        $script:RunErrorsPath = $Run.Errors
        $script:RunMode       = $true
        $OutputPath           = $Run.Csv
        $MetadataPath         = Join-Path $Run.Metadata "ExtensionsBloquees.json"
        $LogPath              = Join-Path $Run.Logs "Get-BlockedExtensions.log"
        $CsvPathConsolide     = Join-Path $Run.Csv "ExtensionsBloquees.csv"
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

    $globalStats = @{
        N1Reussis = 0
        N1Refuses = 0
        N1Erreurs = 0
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

        try {
            $null = Get-Item -Path $cheminN1 -ErrorAction Stop
        }
        catch [System.UnauthorizedAccessException] {
            Write-Log "ACCESS DENIED sur $nomFS - skip" "WARN"
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
            $scanN1 = Invoke-N1BlockedExtensionsScan -CheminN1 $cheminN1 -NomFS $nomFS
            $resultsN1 = $scanN1.Results
            $statsN1 = $scanN1.Stats

            if ($PSCmdlet.ParameterSetName -eq 'Mapping') {
                $csvN1Path = Join-Path -Path $OutputPath -ChildPath "ExtensionsBloquees_${safeNomFS}_${timestamp}.csv"
            }
            else {
                $csvN1Path = $CsvPathConsolide
            }

            if ($resultsN1.Count -gt 0) {
                $resultsN1 | Sort-Object Extension, CheminComplet | Export-Csv -Path $csvN1Path -NoTypeInformation -Delimiter ";" -Encoding UTF8
            }
            else {
                @(New-NoResultRow -NomFileShare $nomFS -Message "Aucun fichier avec extension bloquée détecté") | Export-Csv -Path $csvN1Path -NoTypeInformation -Delimiter ";" -Encoding UTF8
            }
            Write-Log "CSV par N1 généré : $csvN1Path" "SUCCESS"

            $allResults.AddRange($resultsN1)
            $ItemsScanned += [int]$statsN1.ItemsScanned
            $ErrorCount += [int]$statsN1.BlockedCount
            $globalStats.N1Reussis++
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
    $ResultsCount = $allResults.Count

    if ($ResultsCount -gt 0) {
        $allResults | Sort-Object Extension, CheminComplet | Export-Csv -Path $CsvPathConsolide -NoTypeInformation -Delimiter ";" -Encoding UTF8

        $summaryByExt = $allResults | Group-Object Extension | Sort-Object Count -Descending
        Write-Log "Analyse terminée avec $ResultsCount fichier(s) à extension bloquée détecté(s)." "ERROR"
        foreach ($groupe in $summaryByExt) {
            Write-Log ("{0} : {1} fichier(s)" -f $groupe.Name, $groupe.Count) "ERROR"
        }
    }
    else {
        @(New-NoResultRow -NomFileShare "[CONSOLIDE]" -Message "Aucun fichier avec extension bloquée détecté") | Export-Csv -Path $CsvPathConsolide -NoTypeInformation -Delimiter ";" -Encoding UTF8
        Write-Log "Aucun fichier avec extension bloquée détecté." "SUCCESS"
    }

    $mode = if ($PSCmdlet.ParameterSetName -eq 'Mapping') { 'Mapping' } else { 'Single' }
    $durationSeconds = [int][math]::Round(((Get-Date) - $StartTime).TotalSeconds)
    $metadata = [ordered]@{
        session_id       = $timestamp
        vm_source        = $env:COMPUTERNAME
        compte_execution = Get-ExecutionAccount
        mode             = $mode
        mapping_csv      = if ($mode -eq 'Mapping') { $MappingCsv } else { $null }
        n1_total         = $cheminsAScanner.Count
        n1_reussis       = $globalStats.N1Reussis
        n1_refuses       = $globalStats.N1Refuses
        n1_erreurs       = $globalStats.N1Erreurs
        duration_seconds = $durationSeconds
        csv_consolide    = $CsvPathConsolide
    }
    $metadata | ConvertTo-Json -Depth 5 | Out-File -FilePath $MetadataPath -Encoding UTF8

    if ($globalStats.N1Reussis -eq 0 -and $globalStats.N1Refuses -gt 0) {
        $Status = "PartialSuccess"
    }
    if ($ErrorCount -gt 0 -and $globalStats.N1Reussis -eq 0) {
        $Status = "Failed"
    }
    elseif ($ErrorCount -gt 0) {
        $Status = "PartialSuccess"
    }

    $EndTime = Get-Date
    Write-Progress -Activity "Détection des extensions bloquées" -Completed

    $executionChemin = if ($PSCmdlet.ParameterSetName -eq 'Mapping') { "MappingCsv:$MappingCsv" } else { $CheminUNC }
    Write-ExecutionLog -LogPath $LogPath -ScriptName $ScriptName -StartTime $StartTime -EndTime $EndTime -CheminUNC $executionChemin -OutputCsv $CsvPathConsolide -ElementsAnalyses $ItemsScanned -ResultatsTrouves $ResultsCount -WarningCount $WarningCount -ErrorCount $ErrorCount -Status $Status

    Write-Log "Fichier CSV consolidé : $CsvPathConsolide" "SUCCESS"
    Write-Log "Fichier metadata : $MetadataPath" "SUCCESS"
    Write-Log "Fichier LOG : $LogPath" "SUCCESS"
    return $CsvPathConsolide
}