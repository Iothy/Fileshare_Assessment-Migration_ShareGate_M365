<#
.SYNOPSIS
    Détecte les fichiers et dossiers dont le nom est incompatible avec SharePoint Online / OneDrive (mode mono-chemin ou multi-chemins via mapping CSV).

.DESCRIPTION
    Ce script analyse récursivement un ou plusieurs chemins source (FileShare local ou UNC) et identifie
    les fichiers et dossiers présentant des problèmes de nommage pour une migration vers
    Microsoft 365 / SharePoint Online.

    Contrôles réalisés :
    - Caractères invalides : " * : < > ? / \ |
    - Nom commençant par # ou %
    - Nom se terminant par un espace ou un point
    - Noms réservés Windows / SharePoint
    - Présence de _vti_
    - Nom "forms"
    - Nom "desktop.ini"
    - Nom ".lock"
    - Nom commençant par ~$ (fichiers temporaires Office)
    - Nom contenant .. (double point)
    - Nom dépassant 255 caractères

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
    .\Get-InvalidCharacters.ps1 -CheminUNC "\\serveur\partage"

.EXAMPLE
    # Mode multi-chemins (production)
    .\Get-InvalidCharacters.ps1 -MappingCsv ".\Config\FileShareMapping.csv"

.EXAMPLE
    # Exécution avec compte de service (runas /netonly)
    runas /netonly /user:DOMAINE\COMPTE_SERVICE powershell.exe
    # Dans la nouvelle fenêtre PowerShell :
    .\Get-InvalidCharacters.ps1 -MappingCsv ".\Config\FileShareMapping.csv"

.NOTES
    Projet  : Carambar - Migration FileShare vers M365
    Phase   : 01 - Assessment
    Version : v1.1 — Reclassification INFO + suppression règle DoublePoint (sources officielles MS)

    Sources de référence :
    - MS Support — Restrictions and limitations in OneDrive and SharePoint :
      https://support.microsoft.com/en-us/office/restrictions-and-limitations-in-onedrive-and-sharepoint-64883a5d-228e-48f5-b3d2-eb39e07630fa
    - MS Learn — SharePoint Online limits :
      https://learn.microsoft.com/en-us/office365/servicedescriptions/sharepoint-online-service-description/sharepoint-online-limits
    - ShareGate — Special and illegal characters :
      https://help.sharegate.com/en/articles/10236480-special-and-illegal-characters

    Règles de sévérité recalibrées :
    - ERROR  : caractères réellement bloquants (MS doc officielle)
    - WARN   : peut migrer mais avec risque/limitation
    - INFO   : fichiers skippés nativement par ShareGate (aucune action requise)

    Règle DoublePoint supprimée : '..' dans un nom de fichier ne correspond à aucune restriction
    documentée dans les sources MS officielles pour SharePoint Online.

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
    [PSCustomObject]$Run = $null
)

begin {
    Import-Module (Join-Path -Path $PSScriptRoot -ChildPath "Modules\Carambar.Assessment.psm1") -Force
    $outputModulePath = Join-Path -Path $PSScriptRoot -ChildPath "Modules\Carambar.Output.psm1"
    if (Test-Path $outputModulePath) { Import-Module $outputModulePath -Force }

    $ScriptName = Split-Path -Leaf $PSCommandPath
    $ScriptBaseName = [System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)
    $StartTime = Get-Date
    $WarningCount = 0
    $ErrorCount = 0
    $ItemsScanned = 0
    $Status = "Success"

    # Mapping type de problème → sévérité ShareGate
    # ERROR = l'item ne migrera PAS (correction obligatoire, source MS officielle)
    # WARN  = l'item migrera mais avec un risque/limitation
    # INFO  = fichier skippé nativement par ShareGate (aucune action requise)
    $severiteParType = @{
        "CaractereInvalide"     = "ERROR"   # ShareGate bloque ces items
        "PrefixeInterdit"       = "WARN"    # Dépend de la config tenant
        "SuffixeInterdit"       = "ERROR"   # ShareGate bloque ces items
        "NomReserve"            = "ERROR"   # Noms réservés Windows / SharePoint
        "NomReserveSharePoint"  = "ERROR"   # _vti_, forms — réservés SharePoint
        "FichierSysteme"        = "INFO"    # desktop.ini — skippé automatiquement par ShareGate
        "FichierVerrouillage"   = "INFO"    # .lock — skippé automatiquement par ShareGate
        "FichierTemporaire"     = "INFO"    # ~$ — skippé automatiquement par ShareGate
        "NomTropLong"           = "ERROR"   # Nom > 255 car. bloquant pour SPO
    }

    # Mapping type de problème → texte ResultatAnalyse
    $resultatParType = @{
        "CaractereInvalide"     = "Anomalie détectée — renommage requis"
        "PrefixeInterdit"       = "Anomalie détectée — préfixe non supporté"
        "SuffixeInterdit"       = "Anomalie détectée — suffixe non supporté"
        "NomReserve"            = "Anomalie détectée — nom réservé"
        "NomReserveSharePoint"  = "Anomalie détectée — nom réservé SharePoint"
        "FichierSysteme"        = "Fichier système Windows — skippé automatiquement par ShareGate (aucune action)"
        "FichierVerrouillage"   = "Fichier de verrouillage — skippé automatiquement par ShareGate (aucune action)"
        "FichierTemporaire"     = "Fichier temporaire Office — skippé automatiquement par ShareGate (aucune action)"
        "NomTropLong"           = "Anomalie détectée — nom trop long"
    }

    function Add-Issue {
        param(
            [System.Collections.Generic.List[object]]$Collection,
            [System.IO.FileSystemInfo]$Item,
            [string]$ProblemType,
            [string]$Detail,
            [string]$NomFS
        )

        $severite  = if ($severiteParType.ContainsKey($ProblemType)) { $severiteParType[$ProblemType] } else { "WARN" }
        $resultat  = if ($resultatParType.ContainsKey($ProblemType))  { $resultatParType[$ProblemType]  } else { "Anomalie détectée" }

        $Collection.Add([PSCustomObject]@{
                CheminComplet   = $Item.FullName
                Nom             = $Item.Name
                TypeElement     = if ($Item.PSIsContainer) { "Dossier" } else { "Fichier" }
                ProblemeType    = $ProblemType
                Detail          = $Detail
                Severite        = $severite
                LongueurNom     = $Item.Name.Length
                DateAnalyse     = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                ResultatAnalyse = $resultat
                NomFileShare    = $NomFS
            })
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
            Nom             = ""
            TypeElement     = ""
            ProblemeType    = ""
            Detail          = ""
            Severite        = ""
            LongueurNom     = ""
            DateAnalyse     = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            ResultatAnalyse = $Message
            NomFileShare    = $NomFileShare
        }
    }

    function Invoke-N1InvalidCharactersScan {
        param(
            [Parameter(Mandatory = $true)][string]$CheminN1,
            [Parameter(Mandatory = $true)][string]$NomFS
        )

        $results = New-Object 'System.Collections.Generic.List[object]'
        $itemsScannedLocal = 0
        $scanErrorsList = New-Object 'System.Collections.Generic.List[object]'
        $safeNomFS = Get-SafeFileName -Value $NomFS

        $invalidCharsPattern = '[\"*:<>?/\\|]'
        $reservedNames = @(
            "CON", "PRN", "AUX", "NUL",
            "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9",
            "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9",
            "_VTI_", "DESKTOP.INI", ".LOCK", "FORMS"
        )

        Write-Log "Début de l'analyse du chemin : $CheminN1"

        # Énumération résiliente — aucun crash possible sur objets inaccessibles
        Invoke-SafeRecursiveScan -RootPath $CheminN1 -ErrorCollection $scanErrorsList `
            -NomFileShare $NomFS -ProgressActivity "Analyse des caractères invalides" | ForEach-Object {
            $itemsScannedLocal++
            $item = $_

            Write-Verbose "Analyse [$NomFS][$itemsScannedLocal] : $($item.FullName)"

            # Wrapper obligatoire : tout accès aux propriétés en try/catch individuel
            try {
                $name = $item.Name
                $nameWithoutExtension = [System.IO.Path]::GetFileNameWithoutExtension($name)
                $upperName = $name.ToUpperInvariant()
                $upperBaseName = $nameWithoutExtension.ToUpperInvariant()

                if ($name -match $invalidCharsPattern) {
                    Add-Issue -Collection $results -Item $item -ProblemType "CaractereInvalide" -Detail "Le nom contient un ou plusieurs caractères invalides pour SharePoint Online / OneDrive." -NomFS $NomFS
                }

                if ($name.StartsWith("#") -or $name.StartsWith("%")) {
                    Add-Issue -Collection $results -Item $item -ProblemType "PrefixeInterdit" -Detail "Le nom commence par # ou %." -NomFS $NomFS
                }

                if ($name -match '[\. ]$') {
                    Add-Issue -Collection $results -Item $item -ProblemType "SuffixeInterdit" -Detail "Le nom se termine par un espace ou un point." -NomFS $NomFS
                }

                if ($reservedNames -contains $upperName -or $reservedNames -contains $upperBaseName) {
                    Add-Issue -Collection $results -Item $item -ProblemType "NomReserve" -Detail "Le nom est réservé ou non recommandé pour Windows / SharePoint." -NomFS $NomFS
                }

                if ($upperName -like "*_VTI_*") {
                    Add-Issue -Collection $results -Item $item -ProblemType "NomReserveSharePoint" -Detail "Le nom contient _vti_, réservé par SharePoint." -NomFS $NomFS
                }

                if ($upperName -eq "FORMS") {
                    Add-Issue -Collection $results -Item $item -ProblemType "NomReserveSharePoint" -Detail "Le nom 'forms' est réservé dans SharePoint." -NomFS $NomFS
                }

                if ($upperName -eq "DESKTOP.INI") {
                    Add-Issue -Collection $results -Item $item -ProblemType "FichierSysteme" -Detail "Fichier système Windows — skippé automatiquement par ShareGate, aucune action requise." -NomFS $NomFS
                }

                if ($upperName -eq ".LOCK") {
                    Add-Issue -Collection $results -Item $item -ProblemType "FichierVerrouillage" -Detail "Fichier de verrouillage — skippé automatiquement par ShareGate, aucune action requise." -NomFS $NomFS
                }

                if ($name -match '^~\$') {
                    Add-Issue -Collection $results -Item $item -ProblemType "FichierTemporaire" -Detail "Fichier temporaire Office (~$) — skippé automatiquement par ShareGate, aucune action requise." -NomFS $NomFS
                }

                if ($name.Length -gt 255) {
                    Add-Issue -Collection $results -Item $item -ProblemType "NomTropLong" -Detail "Le nom du fichier dépasse 255 caractères (limite SharePoint Online)." -NomFS $NomFS
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

        Write-Log "Nombre d'éléments analysés ($NomFS) : $itemsScannedLocal"

        return @{
            Results = $results
            Stats   = @{
                ItemsScanned = $itemsScannedLocal
            }
        }
    }

    $script:RunErrorsPath = $null
    $script:RunMode       = $false
    Write-Log "Initialisation de l'analyse des caractères invalides"
    Initialize-OutputPath -Path $OutputPath

    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $CsvPathConsolide = Join-Path -Path $OutputPath -ChildPath "CaracteresInvalides_$timestamp.csv"
    $MetadataPath = Join-Path -Path $OutputPath -ChildPath "CaracteresInvalides_Metadata_$timestamp.json"
    $LogPath = Join-Path -Path $OutputPath -ChildPath "$ScriptBaseName`_$timestamp.log"
    Set-LogFile -Path $LogPath
    # Routing: si $Run fourni → chemins hiérarchiques ; sinon → mode legacy
    if ($null -ne $Run) {
        $script:RunErrorsPath = $Run.Errors
        $script:RunMode       = $true
        $OutputPath           = $Run.Csv
        $MetadataPath         = Join-Path $Run.Metadata "CaracteresInvalides.json"
        $LogPath              = Join-Path $Run.Logs "Get-InvalidCharacters.log"
        $CsvPathConsolide     = Join-Path $Run.Csv "CaracteresInvalides.csv"
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
            $scanN1 = Invoke-N1InvalidCharactersScan -CheminN1 $cheminN1 -NomFS $nomFS
            $resultsN1 = $scanN1.Results
            $statsN1 = $scanN1.Stats

            if ($PSCmdlet.ParameterSetName -eq 'Mapping') {
                $csvN1Path = Join-Path -Path $OutputPath -ChildPath "CaracteresInvalides_${safeNomFS}_${timestamp}.csv"
            }
            else {
                $csvN1Path = $CsvPathConsolide
            }

            if ($resultsN1.Count -gt 0) {
                $resultsN1 | Sort-Object CheminComplet, ProblemeType | Export-Csv -Path $csvN1Path -NoTypeInformation -Delimiter ";" -Encoding UTF8
            }
            else {
                @(New-NoResultRow -NomFileShare $nomFS -Message "Aucune anomalie détectée") | Export-Csv -Path $csvN1Path -NoTypeInformation -Delimiter ";" -Encoding UTF8
            }
            Write-Log "CSV par N1 généré : $csvN1Path" "SUCCESS"

            $allResults.AddRange($resultsN1)
            $ItemsScanned += [int]$statsN1.ItemsScanned
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
    $ResultsCount = $allResults.Count

    if ($ResultsCount -gt 0) {
        $allResults | Sort-Object CheminComplet, ProblemeType | Export-Csv -Path $CsvPathConsolide -NoTypeInformation -Delimiter ";" -Encoding UTF8

        $summary = $allResults | Group-Object ProblemeType | Sort-Object Count -Descending
        $erreurCount = ($allResults | Where-Object { $_.Severite -eq "ERROR" } | Measure-Object).Count
        $avertissementCount = ($allResults | Where-Object { $_.Severite -eq "WARN" } | Measure-Object).Count
        Write-Log "Analyse terminée avec $ResultsCount problème(s) détecté(s) ($erreurCount ERROR, $avertissementCount WARN)." "ERROR"
        $ErrorCount++
        foreach ($group in $summary) {
            $niveauGroupe = if ($severiteParType.ContainsKey($group.Name)) { $severiteParType[$group.Name] } else { "WARN" }
            Write-Log ("{0} : {1}" -f $group.Name, $group.Count) $niveauGroupe
            if ($niveauGroupe -eq "ERROR") { $ErrorCount++ } else { $WarningCount++ }
        }
    }
    else {
        @(New-NoResultRow -NomFileShare "[CONSOLIDE]" -Message "Aucune anomalie détectée") | Export-Csv -Path $CsvPathConsolide -NoTypeInformation -Delimiter ";" -Encoding UTF8
        Write-Log "Aucune anomalie de nommage détectée." "SUCCESS"
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
    Write-Progress -Activity "Analyse des caractères invalides" -Completed

    $executionChemin = if ($PSCmdlet.ParameterSetName -eq 'Mapping') { "MappingCsv:$MappingCsv" } else { $CheminUNC }
    Write-ExecutionLog -LogPath $LogPath -ScriptName $ScriptName -StartTime $StartTime -EndTime $EndTime -CheminUNC $executionChemin -OutputCsv $CsvPathConsolide -ElementsAnalyses $ItemsScanned -ResultatsTrouves $ResultsCount -WarningCount $WarningCount -ErrorCount $ErrorCount -Status $Status

    Write-Log "Fichier CSV consolidé : $CsvPathConsolide" "SUCCESS"
    Write-Log "Fichier metadata : $MetadataPath" "SUCCESS"
    Write-Log "Fichier LOG : $LogPath" "SUCCESS"
    return $CsvPathConsolide
}
