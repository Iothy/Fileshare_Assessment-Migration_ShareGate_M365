<#
.SYNOPSIS
    Détecte les fichiers et dossiers dont le chemin dépasse la longueur maximale autorisée pour la migration vers M365 (mode mono-chemin ou multi-chemins via mapping CSV).

.DESCRIPTION
    Ce script analyse récursivement un ou plusieurs chemins source (local ou UNC) et identifie
    les fichiers et dossiers dont le chemin complet dépasse une longueur donnée.

    Ce contrôle est utile dans le cadre d'une migration vers SharePoint Online,
    OneDrive ou Microsoft 365, où les limites de longueur peuvent bloquer ou
    compliquer certaines migrations.

    Deux modes d'utilisation :
    - Mode Single : -CheminUNC <chemin> (rétrocompatible)
    - Mode Mapping : -MappingCsv <chemin du fichier FileShareMapping.csv>

.PARAMETER CheminUNC
    Chemin local ou UNC à analyser (mode Single).

.PARAMETER MappingCsv
    Chemin du fichier mapping CSV (mode Mapping).
    Le CSV doit contenir exactement : SourcePath;TargetType;TargetSPOURL;TargetFolder;DateFilter (YYYY-DD-MM);Permissions

.PARAMETER SpoPathLimit
    Longueur maximale du chemin cible SharePoint Online décodé complet, nom du fichier inclus.
    Valeur par défaut : 400

.PARAMETER WindowsOfficePathLimit
    Limite projet pour la compatibilité Windows/Office locale.
    Valeur par défaut : 256

.PARAMETER EstimatedLocalPrefixLength
    Longueur estimée du préfixe local OneDrive. Hypothèse projet configurable, pas une limite Microsoft.
    Valeur par défaut : 96

.PARAMETER OutputPath
    Dossier de sortie des résultats.
    Par défaut : .\Output (dans le dossier du script)

.EXAMPLE
    # Mode mono-chemin (rétrocompatible)
    .\Get-PathTooLong.ps1 -CheminUNC "\\serveur\partage" -SpoPathLimit 400

.EXAMPLE
    # Mode multi-chemins (production)
    .\Get-PathTooLong.ps1 -MappingCsv ".\Config\FileShareMapping.PrimaGAZ.csv" -SpoPathLimit 400

.EXAMPLE
    # Mode mono-chemin avec limites configurables
    .\Get-PathTooLong.ps1 -CheminUNC "\\serveur\partage" -SpoPathLimit 400 -WindowsOfficePathLimit 256 -EstimatedLocalPrefixLength 96

.NOTES
    Projet  : PrimaGAZ - Migration FileShare vers M365
    Phase   : 01 - Assessment

    Authentification :
    Pour utiliser un compte de service différent de la session courante,
    lancer une session PowerShell via : runas /netonly /user:DOMAIN\user powershell.exe
#>

[CmdletBinding(DefaultParameterSetName = 'Single')]
param(
    [Parameter(Mandatory, ParameterSetName = 'Single', HelpMessage = "Chemin local ou UNC à analyser")]
    [ValidateNotNullOrEmpty()]
    [string]$CheminUNC,

    [Parameter(Mandatory = $false, ParameterSetName = 'Single', HelpMessage = "URL cible SharePoint Online pour simuler les chemins en mode mono-chemin")]
    [AllowEmptyString()]
    [string]$TargetSPOURL = '',

    [Parameter(Mandatory = $false, ParameterSetName = 'Single', HelpMessage = "Type de destination pour simuler les chemins en mode mono-chemin")]
    [AllowEmptyString()]
    [string]$TargetType = '',

    [Parameter(Mandatory = $false, ParameterSetName = 'Single', HelpMessage = "Dossier cible pour simuler les chemins en mode mono-chemin")]
    [AllowEmptyString()]
    [string]$TargetFolder = '',

    [Parameter(Mandatory, ParameterSetName = 'Mapping', HelpMessage = "Chemin du fichier FileShareMapping.csv (colonnes : SourcePath;TargetType;TargetSPOURL;TargetFolder;DateFilter (YYYY-DD-MM);Permissions)")]
    [ValidateNotNullOrEmpty()]
    [string]$MappingCsv,

    [Parameter(Mandatory = $false, HelpMessage = "Limite SharePoint Online du chemin cible décodé complet")]
    [Alias('LongueurMax')]
    [ValidateRange(1, 32767)]
    [int]$SpoPathLimit = 400,

    [Parameter(Mandatory = $false, HelpMessage = "Limite projet Windows/Office locale")]
    [ValidateRange(1, 32767)]
    [int]$WindowsOfficePathLimit = 256,

    [Parameter(Mandatory = $false, HelpMessage = "Longueur estimée du préfixe local OneDrive")]
    [Alias('UrlPrefixLength')]
    [ValidateRange(0, 32767)]
    [int]$EstimatedLocalPrefixLength = 96,

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
    $fileShareAssessmentModulePath = Join-Path -Path $PSScriptRoot -ChildPath 'FileShareAssessment/FileShareAssessment.psd1'
    if (Test-Path $fileShareAssessmentModulePath) { Import-Module $fileShareAssessmentModulePath -Force }
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

    function New-NoResultRow {
        param(
            [string]$NomFileShare,
            [string]$SourcePath,
            [string]$TargetSPOURL,
            [string]$TargetType,
            [string]$TargetFolder,
            [string]$Status,
            [string]$Message
        )

        return [PSCustomObject]@{
            SourcePath                    = $SourcePath
            SourceRelativePath            = ''
            TargetSPOURL                  = $TargetSPOURL
            TargetType                    = $TargetType
            TargetFolder                  = $TargetFolder
            SimulatedTargetPath           = ''
            TypeElement                   = ''
            Nom                           = ''
            LongueurSource                = ''
            LongueurCibleSPO              = ''
            LimiteSPO                     = $SpoPathLimit
            DepassementSPO                = ''
            CompatibleSPO                 = ''
            LongueurRelativeWindowsOffice = ''
            BudgetWindowsOffice           = [Math]::Max(0, $WindowsOfficePathLimit - $EstimatedLocalPrefixLength)
            DepassementWindowsOffice      = ''
            CompatibleWindowsOffice       = ''
            StatutControle                = $Status
            ActionRecommandee             = $Message
            DateAnalyse                   = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            NomFileShare                  = $NomFileShare
        }
    }

    function Get-PathTooLongColumns {
        return @(
            'SourcePath',
            'SourceRelativePath',
            'TargetSPOURL',
            'TargetType',
            'TargetFolder',
            'SimulatedTargetPath',
            'TypeElement',
            'Nom',
            'LongueurSource',
            'LongueurCibleSPO',
            'LimiteSPO',
            'DepassementSPO',
            'CompatibleSPO',
            'LongueurRelativeWindowsOffice',
            'BudgetWindowsOffice',
            'DepassementWindowsOffice',
            'CompatibleWindowsOffice',
            'StatutControle',
            'ActionRecommandee',
            'DateAnalyse',
            'NomFileShare'
        )
    }

    function Join-UrlPath {
        param([string[]]$Parts)

        $segments = @(
            foreach ($part in $Parts) {
                if (-not [string]::IsNullOrWhiteSpace($part)) {
                    $part.Trim().Replace('\', '/').Trim('/')
                }
            }
        ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

        return ($segments -join '/')
    }

    function Get-DecodedLength {
        param([string]$Value)

        if ([string]::IsNullOrEmpty($Value)) { return 0 }
        try {
            return ([System.Uri]::UnescapeDataString($Value)).Length
        }
        catch {
            return $Value.Length
        }
    }

    function Invoke-N1PathTooLongScan {
        param(
            [Parameter(Mandatory = $true)][string]$CheminN1,
            [Parameter(Mandatory = $true)][string]$NomFS,
            [Parameter(Mandatory = $true)][string]$TargetSPOURL,
            [AllowEmptyString()][string]$TargetType,
            [AllowEmptyString()][string]$TargetFolder,
            [AllowEmptyString()][string]$LeafName,
            [Parameter(Mandatory = $true)][int]$SpoPathLimit,
            [Parameter(Mandatory = $true)][int]$WindowsOfficePathLimit,
            [Parameter(Mandatory = $true)][int]$EstimatedLocalPrefixLength
        )

        $results = New-Object 'System.Collections.Generic.List[object]'
        $itemsScannedLocal = 0
        $scanErrorsList = New-Object 'System.Collections.Generic.List[object]'
        $safeNomFS = Get-SafeFileName -Value $NomFS

        Write-Log "Début de l'analyse du chemin : $CheminN1"
        Write-Log "Limite SharePoint Online appliquée : $SpoPathLimit caractères"
        Write-Log "Limite Windows/Office projet : $WindowsOfficePathLimit caractères ; préfixe local estimé : $EstimatedLocalPrefixLength ; budget relatif : $([Math]::Max(0, $WindowsOfficePathLimit - $EstimatedLocalPrefixLength)) caractères"

        # Énumération résiliente — aucun crash possible sur objets inaccessibles
        Invoke-SafeRecursiveScan -RootPath $CheminN1 -ErrorCollection $scanErrorsList `
            -NomFileShare $NomFS -ProgressActivity "Analyse des chemins trop longs" | ForEach-Object {
            $itemsScannedLocal++
            $item = $_

            Write-Verbose "Analyse [$NomFS][$itemsScannedLocal] : $($item.FullName)"

            # Wrapper obligatoire : tout accès aux propriétés en try/catch individuel
            try {
                $currentLength = $item.FullName.Length
                $rootLength = $CheminN1.TrimEnd('\').Length
                $relativePath = if ($currentLength -gt $rootLength) { $item.FullName.Substring($rootLength).TrimStart('\') } else { "" }
                $targetRelativePath = Join-UrlPath -Parts @($TargetFolder, $LeafName, $relativePath)
                $simulatedTargetPath = Join-UrlPath -Parts @($TargetSPOURL, $targetRelativePath)
                $spoLength = Get-DecodedLength -Value $simulatedTargetPath
                $windowsOfficeLength = Get-DecodedLength -Value $targetRelativePath
                $windowsOfficeBudget = [Math]::Max(0, $WindowsOfficePathLimit - $EstimatedLocalPrefixLength)
                $isDir = ($item -is [System.IO.DirectoryInfo])
                $compatibleSPO = $spoLength -le $SpoPathLimit
                $compatibleWindowsOffice = $windowsOfficeLength -le $windowsOfficeBudget

                if (-not $compatibleSPO -or -not $compatibleWindowsOffice) {
                    $actions = @()
                    if (-not $compatibleSPO) { $actions += 'Réduire le chemin cible SharePoint Online.' }
                    if (-not $compatibleWindowsOffice) { $actions += 'Réduire le chemin relatif synchronisé pour la compatibilité Windows/Office.' }
                    $results.Add([PSCustomObject]@{
                            SourcePath                    = $item.FullName
                            SourceRelativePath            = $relativePath
                            TargetSPOURL                  = $TargetSPOURL
                            TargetType                    = $TargetType
                            TargetFolder                  = $TargetFolder
                            SimulatedTargetPath           = $simulatedTargetPath
                            TypeElement                   = if ($isDir) { "Dossier" } else { "Fichier" }
                            Nom                           = $item.Name
                            LongueurSource                = $currentLength
                            LongueurCibleSPO              = $spoLength
                            LimiteSPO                     = $SpoPathLimit
                            DepassementSPO                = [Math]::Max(0, $spoLength - $SpoPathLimit)
                            CompatibleSPO                 = $compatibleSPO
                            LongueurRelativeWindowsOffice = $windowsOfficeLength
                            BudgetWindowsOffice           = $windowsOfficeBudget
                            DepassementWindowsOffice      = [Math]::Max(0, $windowsOfficeLength - $windowsOfficeBudget)
                            CompatibleWindowsOffice       = $compatibleWindowsOffice
                            StatutControle                = 'Completed'
                            ActionRecommandee             = ($actions -join ' ')
                            DateAnalyse                   = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                            NomFileShare                  = $NomFS
                        })
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
                $deniedCsv = if ($Run -and -not [string]::IsNullOrWhiteSpace($sourceIdentifier)) {
                    Get-AssessmentSourceFilePath -Run $Run -SourceIdentifier $sourceIdentifier -Prefix 'AccesRefuses' -Extension 'csv'
                }
                else {
                    Join-Path $errBase "AccessDenied_${safeNomFS}${errSuffix}.csv"
                }
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
    Write-Log "Initialisation de l'analyse des chemins trop longs"
    Initialize-OutputPath -Path $OutputPath

    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $CsvPathConsolide = Join-Path -Path $OutputPath -ChildPath "CheminsTropLongs_$timestamp.csv"
    $MetadataPath = Join-Path -Path $OutputPath -ChildPath "CheminsTropLongs_Metadata_$timestamp.json"
    $LogPath = Join-Path -Path $OutputPath -ChildPath "$ScriptBaseName`_$timestamp.log"
    Set-LogFile -Path $LogPath
    # Routing: si $Run fourni → chemins hiérarchiques ; sinon → mode legacy
    if ($null -ne $Run) {
        $script:RunErrorsPath = $Run.Errors
        $script:RunMode       = $true
        $OutputPath           = $Run.Csv
        $MetadataPath         = Join-Path $Run.Metadata "CheminsTropLongs.json"
        $LogPath              = Join-Path $Run.Logs "Get-PathTooLong.log"
        $CsvPathConsolide     = Join-Path $Run.Csv "CheminsTropLongs.csv"
        Set-LogFile -Path $LogPath
    }
    Write-Log "=== Démarrage $ScriptName ===" "INFO"
    Write-Log "VM source : $env:COMPUTERNAME" "INFO"
    Write-Log "Compte d'exécution : $(Get-ExecutionAccount)" "INFO"
    Write-Log "OutputPath : $OutputPath" "INFO"

    $cheminsAScanner = New-Object 'System.Collections.Generic.List[object]'

    if ($PSCmdlet.ParameterSetName -eq 'Mapping') {
        $mappingValidation = Test-FileShareMapping -Path $MappingCsv
        if (-not $mappingValidation.IsValid) {
            throw (($mappingValidation.Errors | ForEach-Object { 'Ligne {0}: {1}' -f $_.LineNumber, $_.Message }) -join [Environment]::NewLine)
        }

        foreach ($warning in $mappingValidation.Warnings) {
            Write-Log ('Mapping ligne {0}: {1}' -f $warning.LineNumber, $warning.Message) 'WARN'
            $WarningCount++
        }

        foreach ($row in (Import-FileShareMapping -Path $MappingCsv)) {
            $cheminsAScanner.Add($row)
        }

        if ($cheminsAScanner.Count -eq 0) {
            throw "Aucune entrée exploitable trouvée dans '$MappingCsv'."
        }
    }
    else {
        Test-SourcePath -Path $CheminUNC
        $sourceContext = $null
        try {
            $sourceContext = Resolve-FileShareSourceMetadata -SourcePath $CheminUNC
        }
        catch {
            $sourceContext = [PSCustomObject]@{
                SourcePath       = $CheminUNC
                SourceIdentifier = Get-SafeFileName -Value (Get-NomFileShareFromPath -Path $CheminUNC)
            }
        }
        $cheminsAScanner.Add([PSCustomObject]@{
            SourcePath       = $sourceContext.SourcePath
            SourceIdentifier = $sourceContext.SourceIdentifier
            TargetSPOURL     = $TargetSPOURL.TrimEnd('/')
            TargetType       = $TargetType
            TargetFolder     = $TargetFolder.Trim().Replace('\', '/').Trim('/')
            LeafName         = if ($sourceContext.PSObject.Properties.Name -contains 'LeafName') { $sourceContext.LeafName } else { Get-NomFileShareFromPath -Path $CheminUNC }
            LineNumber       = 0
        })
    }

    $allResults = New-Object 'System.Collections.Generic.List[object]'

    $globalStats = @{
        N1Reussis = 0
        N1Refuses = 0
        N1Erreurs = 0
        N1Skipped = 0
    }
}

process {
    if ($Credential) {
        $smbResult = Connect-PrimaGazFileShare -Server $Server -Credential $Credential
        Test-FileShareIdentity -Server $Server -ExpectedUserName $Credential.UserName
        $script:smbConnected = -not $smbResult.AlreadyConnected
    }

    Write-Log "Mode d'exécution : $($PSCmdlet.ParameterSetName)"

    foreach ($entry in $cheminsAScanner) {
        $cheminN1 = if ($entry.PSObject.Properties.Name -contains 'SourcePath') { $entry.SourcePath } else { $entry.CheminUNC }
        if ($entry.PSObject.Properties.Name -contains 'SourceIdentifier' -and -not [string]::IsNullOrWhiteSpace($entry.SourceIdentifier)) {
            $sourceIdentifier = $entry.SourceIdentifier
        }
        elseif ($entry.PSObject.Properties.Name -contains 'NomFileShare' -and -not [string]::IsNullOrWhiteSpace($entry.NomFileShare)) {
            $sourceIdentifier = Get-SafeFileName -Value $entry.NomFileShare
        }
        else {
            $sourceIdentifier = Get-SafeFileName -Value (Get-NomFileShareFromPath -Path $cheminN1)
        }
        $nomFS = $sourceIdentifier
        $safeNomFS = Get-SafeFileName -Value $nomFS
        $targetSpoUrl = if ($entry.PSObject.Properties.Name -contains 'TargetSPOURL') { [string]$entry.TargetSPOURL } else { '' }
        $targetType = if ($entry.PSObject.Properties.Name -contains 'TargetType') { [string]$entry.TargetType } else { '' }
        $targetFolder = if ($entry.PSObject.Properties.Name -contains 'TargetFolder') { [string]$entry.TargetFolder } else { '' }
        $leafName = if ($entry.PSObject.Properties.Name -contains 'LeafName' -and -not [string]::IsNullOrWhiteSpace($entry.LeafName)) {
            [string]$entry.LeafName
        }
        else {
            Get-NomFileShareFromPath -Path $cheminN1
        }
        if ($Run) {
            $csvN1Path = Get-AssessmentSourceFilePath -Run $Run -SourceIdentifier $sourceIdentifier -Prefix 'CheminsLongs' -Extension 'csv'
        }
        elseif ($PSCmdlet.ParameterSetName -eq 'Mapping') {
            $csvN1Path = Join-Path -Path $OutputPath -ChildPath "CheminsTropLongs_${safeNomFS}_${timestamp}.csv"
        }
        else {
            $csvN1Path = $CsvPathConsolide
        }
        if ($Run) {
            $null = Get-AssessmentSourceFolder -Run $Run -SourceIdentifier $sourceIdentifier
            $accessDeniedPath = Get-AssessmentSourceFilePath -Run $Run -SourceIdentifier $sourceIdentifier -Prefix 'AccesRefuses' -Extension 'csv'
            Write-EmptyCsv -Path $accessDeniedPath -Columns @('NomFileShare','Chemin','TypeErreur','ExceptionType','MessageErreur','DateDetection')
        }

        Write-Log "=== Démarrage scan : $nomFS ($cheminN1) ===" "INFO"

        if ([string]::IsNullOrWhiteSpace($targetSpoUrl)) {
            $message = "Contrôle chemins longs ignoré : TargetSPOURL absent pour cette source."
            Write-Log "$message Source : $cheminN1" "WARN"
            @(New-NoResultRow -NomFileShare $nomFS -SourcePath $cheminN1 -TargetSPOURL $targetSpoUrl -TargetType $targetType -TargetFolder $targetFolder -Status 'Skipped' -Message $message) |
                Export-Csv -Path $csvN1Path -NoTypeInformation -Delimiter ";" -Encoding UTF8
            $WarningCount++
            $globalStats.N1Skipped++
            continue
        }

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
            $scanN1 = Invoke-N1PathTooLongScan -CheminN1 $cheminN1 -NomFS $nomFS -TargetSPOURL $targetSpoUrl -TargetType $targetType -TargetFolder $targetFolder -LeafName $leafName -SpoPathLimit $SpoPathLimit -WindowsOfficePathLimit $WindowsOfficePathLimit -EstimatedLocalPrefixLength $EstimatedLocalPrefixLength
            $resultsN1 = $scanN1.Results
            $statsN1 = $scanN1.Stats

            if ($resultsN1.Count -gt 0) {
                $resultsN1 | Sort-Object LongueurCibleSPO -Descending | Export-Csv -Path $csvN1Path -NoTypeInformation -Delimiter ";" -Encoding UTF8
            }
            else {
                if ($Run) {
                    Write-EmptyCsv -Path $csvN1Path -Columns (Get-PathTooLongColumns)
                }
                else {
                    @(New-NoResultRow -NomFileShare $nomFS -SourcePath $cheminN1 -TargetSPOURL $targetSpoUrl -TargetType $targetType -TargetFolder $targetFolder -Status 'Completed' -Message "Aucun dépassement SPO ou Windows/Office détecté") |
                        Export-Csv -Path $csvN1Path -NoTypeInformation -Delimiter ";" -Encoding UTF8
                }
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
    if ($script:smbConnected) {
        try { Disconnect-PrimaGazFileShare -Server $Server } catch { Write-Verbose "Disconnect-PrimaGazFileShare: $($_.Exception.Message)" }
        $script:smbConnected = $false
    }
    $ResultsCount = $allResults.Count

    if ($ResultsCount -gt 0) {
        $allResults | Sort-Object LongueurCibleSPO -Descending | Export-Csv -Path $CsvPathConsolide -NoTypeInformation -Delimiter ";" -Encoding UTF8

        $maxDetected = ($allResults | Measure-Object -Property LongueurCibleSPO -Maximum).Maximum
        $avgDetected = [math]::Round((($allResults | Measure-Object -Property LongueurCibleSPO -Average).Average), 2)

        Write-Log "Analyse terminée avec $ResultsCount chemin(s) trop long(s)." "ERROR"
        Write-Log "Longueur maximale détectée : $maxDetected caractères" "ERROR"
        Write-Log "Longueur moyenne des chemins en anomalie : $avgDetected caractères" "ERROR"
        $ErrorCount += 3
    }
    else {
        @(New-NoResultRow -NomFileShare "[CONSOLIDE]" -SourcePath '' -TargetSPOURL '' -TargetType '' -TargetFolder '' -Status 'Completed' -Message "Aucun dépassement SPO ou Windows/Office détecté") |
            Export-Csv -Path $CsvPathConsolide -NoTypeInformation -Delimiter ";" -Encoding UTF8
        Write-Log "Aucun chemin ne dépasse les limites SPO ($SpoPathLimit) et Windows/Office (budget relatif $([Math]::Max(0, $WindowsOfficePathLimit - $EstimatedLocalPrefixLength)))." "SUCCESS"
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
        n1_skipped       = $globalStats.N1Skipped
        spo_path_limit   = $SpoPathLimit
        windows_office_path_limit = $WindowsOfficePathLimit
        estimated_local_prefix_length = $EstimatedLocalPrefixLength
        windows_office_relative_budget = [Math]::Max(0, $WindowsOfficePathLimit - $EstimatedLocalPrefixLength)
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
    Write-Progress -Activity "Analyse des chemins trop longs" -Completed

    $executionChemin = if ($PSCmdlet.ParameterSetName -eq 'Mapping') { "MappingCsv:$MappingCsv" } else { $CheminUNC }
    Write-ExecutionLog -LogPath $LogPath -ScriptName $ScriptName -StartTime $StartTime -EndTime $EndTime -CheminUNC $executionChemin -OutputCsv $CsvPathConsolide -ElementsAnalyses $ItemsScanned -ResultatsTrouves $ResultsCount -WarningCount $WarningCount -ErrorCount $ErrorCount -Status $Status

    Write-Log "Fichier CSV consolidé : $CsvPathConsolide" "SUCCESS"
    Write-Log "Fichier metadata : $MetadataPath" "SUCCESS"
    Write-Log "Fichier LOG : $LogPath" "SUCCESS"
    return $CsvPathConsolide
}