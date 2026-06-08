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
    Le CSV doit contenir au minimum les colonnes : CheminUNC;NomFileShare

.PARAMETER LongueurMax
    Longueur maximale autorisée du chemin complet.
    Valeur par défaut : 400

.PARAMETER UrlPrefixLength
    Longueur estimée du préfixe URL SharePoint (ex: https://tenant.sharepoint.com/sites/site/lib/ = ~70 car.).
    Si fourni (> 0), la longueur simulée de l'URL cible est calculée et comparée à LongueurMax.
    Valeur par défaut : 0 (désactivé)

.PARAMETER OutputPath
    Dossier de sortie des résultats.
    Par défaut : .\Output (dans le dossier du script)

.EXAMPLE
    # Mode mono-chemin (rétrocompatible)
    .\Get-PathTooLong.ps1 -CheminUNC "\\serveur\partage" -LongueurMax 400

.EXAMPLE
    # Mode multi-chemins (production)
    .\Get-PathTooLong.ps1 -MappingCsv ".\Config\FileShareMapping.PrimaGAZ.csv" -LongueurMax 400

.EXAMPLE
    # Mode mono-chemin avec simulation URL SharePoint
    .\Get-PathTooLong.ps1 -CheminUNC "\\serveur\partage" -LongueurMax 400 -UrlPrefixLength 72

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

    [Parameter(Mandatory, ParameterSetName = 'Mapping', HelpMessage = "Chemin du fichier FileShareMapping.csv (colonnes : CheminUNC;NomFileShare;...)")]
    [ValidateNotNullOrEmpty()]
    [string]$MappingCsv,

    [Parameter(Mandatory = $false, HelpMessage = "Longueur maximale du chemin complet")]
    [ValidateRange(1, 32767)]
    [int]$LongueurMax = 400,

    [Parameter(Mandatory = $false, HelpMessage = "Longueur estimée du préfixe URL SharePoint (ex: https://tenant.sharepoint.com/sites/site/lib/ = ~70 car.)")]
    [ValidateRange(0, 500)]
    [int]$UrlPrefixLength = 0,

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

    function New-NoResultRow {
        param(
            [string]$NomFileShare,
            [string]$Message
        )

        return [PSCustomObject]@{
            CheminComplet      = ""
            Nom                = ""
            TypeElement        = ""
            LongueurChemin     = ""
            LongueurMax        = $LongueurMax
            Depassement        = ""
            LongueurUrlSimulee = ""
            LongueurNom        = ""
            Severite           = ""
            DateAnalyse        = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            ResultatAnalyse    = $Message
            NomFileShare       = $NomFileShare
        }
    }

    function Invoke-N1PathTooLongScan {
        param(
            [Parameter(Mandatory = $true)][string]$CheminN1,
            [Parameter(Mandatory = $true)][string]$NomFS,
            [Parameter(Mandatory = $true)][int]$LongueurMax,
            [Parameter(Mandatory = $true)][int]$UrlPrefixLength
        )

        $results = New-Object 'System.Collections.Generic.List[object]'
        $itemsScannedLocal = 0
        $scanErrorsList = New-Object 'System.Collections.Generic.List[object]'
        $safeNomFS = Get-SafeFileName -Value $NomFS

        Write-Log "Début de l'analyse du chemin : $CheminN1"
        Write-Log "Longueur maximale appliquée : $LongueurMax caractères"
        if ($UrlPrefixLength -gt 0) {
            Write-Log "Préfixe URL SharePoint simulé : $UrlPrefixLength caractères"
        }

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
                $simulatedUrlLength = $UrlPrefixLength + $relativePath.Length
                $effectiveLength = if ($UrlPrefixLength -gt 0) { $simulatedUrlLength } else { $currentLength }
                $isDir = ($item -is [System.IO.DirectoryInfo])

                if ($effectiveLength -gt $LongueurMax) {
                    $results.Add([PSCustomObject]@{
                            CheminComplet      = $item.FullName
                            Nom                = $item.Name
                            TypeElement        = if ($isDir) { "Dossier" } else { "Fichier" }
                            LongueurChemin     = $currentLength
                            LongueurMax        = $LongueurMax
                            Depassement        = $effectiveLength - $LongueurMax
                            LongueurUrlSimulee = if ($UrlPrefixLength -gt 0) { $simulatedUrlLength } else { "N/A" }
                            LongueurNom        = $item.Name.Length
                            Severite           = "ERROR"
                            DateAnalyse        = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                            ResultatAnalyse    = "Chemin trop long détecté"
                            NomFileShare       = $NomFS
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
        $smbResult = Connect-PrimaGazFileShare -Server $Server -Credential $Credential
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
            $scanN1 = Invoke-N1PathTooLongScan -CheminN1 $cheminN1 -NomFS $nomFS -LongueurMax $LongueurMax -UrlPrefixLength $UrlPrefixLength
            $resultsN1 = $scanN1.Results
            $statsN1 = $scanN1.Stats

            if ($PSCmdlet.ParameterSetName -eq 'Mapping') {
                $csvN1Path = Join-Path -Path $OutputPath -ChildPath "CheminsTropLongs_${safeNomFS}_${timestamp}.csv"
            }
            else {
                $csvN1Path = $CsvPathConsolide
            }

            if ($resultsN1.Count -gt 0) {
                $resultsN1 | Sort-Object LongueurChemin -Descending | Export-Csv -Path $csvN1Path -NoTypeInformation -Delimiter ";" -Encoding UTF8
            }
            else {
                @(New-NoResultRow -NomFileShare $nomFS -Message "Aucun chemin trop long détecté") | Export-Csv -Path $csvN1Path -NoTypeInformation -Delimiter ";" -Encoding UTF8
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
        $allResults | Sort-Object LongueurChemin -Descending | Export-Csv -Path $CsvPathConsolide -NoTypeInformation -Delimiter ";" -Encoding UTF8

        $maxDetected = ($allResults | Measure-Object -Property LongueurChemin -Maximum).Maximum
        $avgDetected = [math]::Round((($allResults | Measure-Object -Property LongueurChemin -Average).Average), 2)

        Write-Log "Analyse terminée avec $ResultsCount chemin(s) trop long(s)." "ERROR"
        Write-Log "Longueur maximale détectée : $maxDetected caractères" "ERROR"
        Write-Log "Longueur moyenne des chemins en anomalie : $avgDetected caractères" "ERROR"
        $ErrorCount += 3
    }
    else {
        @(New-NoResultRow -NomFileShare "[CONSOLIDE]" -Message "Aucun chemin trop long détecté") | Export-Csv -Path $CsvPathConsolide -NoTypeInformation -Delimiter ";" -Encoding UTF8
        Write-Log "Aucun chemin ne dépasse la longueur maximale de $LongueurMax caractères." "SUCCESS"
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
    Write-Progress -Activity "Analyse des chemins trop longs" -Completed

    $executionChemin = if ($PSCmdlet.ParameterSetName -eq 'Mapping') { "MappingCsv:$MappingCsv" } else { $CheminUNC }
    Write-ExecutionLog -LogPath $LogPath -ScriptName $ScriptName -StartTime $StartTime -EndTime $EndTime -CheminUNC $executionChemin -OutputCsv $CsvPathConsolide -ElementsAnalyses $ItemsScanned -ResultatsTrouves $ResultsCount -WarningCount $WarningCount -ErrorCount $ErrorCount -Status $Status

    Write-Log "Fichier CSV consolidé : $CsvPathConsolide" "SUCCESS"
    Write-Log "Fichier metadata : $MetadataPath" "SUCCESS"
    Write-Log "Fichier LOG : $LogPath" "SUCCESS"
    return $CsvPathConsolide
}