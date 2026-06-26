<#
.SYNOPSIS
    Verifie que les URLs de sites SharePoint/Teams listees dans un CSV de mapping existent
    et sont accessibles avec le compte ShareGate destination.

.DESCRIPTION
    Lit un CSV (separateur point-virgule par defaut) contenant au minimum une colonne
    TargetSPOURL et tente une connexion ShareGate (Connect-Site) sur chaque URL unique.
    Produit un rapport console + un rapport CSV (dans le dossier output) avec le statut
    Found / NotFound / AccessDenied / InvalidUrl / Error.

    Les sorties sont isolees dans un sous-dossier dedie au script :
      - Logs   : C:\migrationFactory\logs\<NomDuScript>\
      - Rapport: C:\migrationFactory\output\<NomDuScript>\
    Le log est journalier (date du jour incluse dans le nom du fichier).

.PARAMETER BatchFile
    Chemin complet vers le CSV a valider, OU simple nom de fichier present dans le dossier
    'input' de la migration factory (C:\migrationFactory\input).
    Colonnes attendues (separateur ';') :
      SourcePath, TargetType, TargetSPOURL, TargetFolder, DateFilter, Permissions

.EXAMPLE
    .\01-Test-SPOTargets.ps1 -BatchFile .\SPO-Input_Worker_1_20260625.csv

.EXAMPLE
    .\01-Test-SPOTargets.ps1 -BatchFile SPO-Input_Worker_1_20260625.csv
    # Le fichier est recherche automatiquement dans C:\migrationFactory\input
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [Alias('BatchName')]
    [string]$BatchFile
)

# Racine de la migration factory (en dur)
$FactoryRoot = 'C:\migrationFactory'

# Valeurs par defaut
$Delimiter = ";"

# --- 0. Pre-requis ---------------------------------------------------------
Import-Module Sharegate -ErrorAction Stop

# Identifiant du script (ex: "01-Test-SPOTargets") => sert a nommer les sous-dossiers
$scriptTag = [System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Path)
if ([string]::IsNullOrWhiteSpace($scriptTag)) { $scriptTag = 'Test-SPOTargets' }

# Structure migrationFactory : sous-dossiers dedies par script pour logs & output
$InputDir  = Join-Path $FactoryRoot 'input'
$OutputDir = Join-Path (Join-Path $FactoryRoot 'output') $scriptTag
$LogsDir   = Join-Path (Join-Path $FactoryRoot 'logs')   $scriptTag
foreach ($d in $InputDir, $OutputDir, $LogsDir) {
    if (-not (Test-Path -LiteralPath $d)) { New-Item -ItemType Directory -Path $d -Force | Out-Null }
}

# Resolution du fichier batch : chemin complet ou nom de fichier dans input/
if (-not (Test-Path -LiteralPath $BatchFile)) {
    $candidate = Join-Path $InputDir (Split-Path -Leaf $BatchFile)
    if (Test-Path -LiteralPath $candidate) {
        $BatchFile = $candidate
    } else {
        throw "Fichier batch introuvable : '$BatchFile' (cherche aussi dans '$InputDir')."
    }
}
$BatchFile = (Resolve-Path -LiteralPath $BatchFile).Path
$batchBase = [System.IO.Path]::GetFileNameWithoutExtension($BatchFile)

# Log journalier : PrimagaZ_Test_Targets_M365_YYYYMMDD.log (date du jour)
$dayTag  = Get-Date -Format 'yyyyMMdd'
$LogFile = Join-Path $LogsDir ("PrimagaZ_Test_Targets_M365_{0}.log" -f $dayTag)

# --- Module de log (fallback inline, meme signature que Avanade.MF.PR.psm1) ---
function Write-Log {
    param($Name, $Msg, $Type = "INFO", $Logpath, $Delimiter = ";")
    $ts    = Get-Date -Format "yyyyMMdd_HHmmss"
    $line  = "$ts ; $Type ; $Name ; $Msg"
    $color = switch ($Type) { "ERROR" { "Red" } "WARN" { "Yellow" } default { "Green" } }
    Write-Host $line -ForegroundColor $color
    if ($Logpath) {
        try {
            $fs = [System.IO.File]::Open($Logpath, [System.IO.FileMode]::Append, [System.IO.FileAccess]::Write, [System.IO.FileShare]::ReadWrite)
            $sw = [System.IO.StreamWriter]::new($fs, [System.Text.Encoding]::UTF8)
            $sw.WriteLine($line); $sw.Flush(); $sw.Close(); $fs.Close()
        } catch {
            Start-Sleep -Milliseconds (Get-Random -Minimum 50 -Maximum 200)
            $line | Out-File -FilePath $Logpath -Append -Encoding UTF8 -ErrorAction SilentlyContinue
        }
    }
}
$log_params = @{ Name = "Test-SPOTargets" ; Logpath = $LogFile }

Write-Log @log_params -Msg "=== Demarrage Test-SPOTargets ===" -Type INFO
Write-Log @log_params -Msg "Batch       : $BatchFile"  -Type INFO
Write-Log @log_params -Msg "Delimiteur  : '$Delimiter'" -Type INFO
Write-Log @log_params -Msg "Dossier log : $LogsDir"   -Type INFO
Write-Log @log_params -Msg "Dossier out : $OutputDir" -Type INFO
Write-Log @log_params -Msg "Log         : $LogFile"   -Type INFO

# --- 1. Lecture & extraction des URLs uniques ------------------------------
$rows = Import-Csv -Path $BatchFile -Delimiter $Delimiter

if (-not ($rows | Get-Member -Name TargetSPOURL -MemberType NoteProperty)) {
    Write-Log @log_params -Msg "Colonne 'TargetSPOURL' absente du fichier '$BatchFile' (verifier le delimiteur)." -Type ERROR
    throw "Colonne 'TargetSPOURL' absente du fichier '$BatchFile' (verifier le delimiteur)."
}

$urls = $rows |
    Where-Object { "$($_.TargetSPOURL)".Trim() } |
    ForEach-Object { "$($_.TargetSPOURL)".Trim() } |
    Sort-Object -Unique

Write-Log @log_params -Msg "$($urls.Count) URL(s) unique(s) a tester (sur $($rows.Count) ligne(s))" -Type INFO

if (-not $urls -or $urls.Count -eq 0) {
    Write-Log @log_params -Msg "Aucune URL a tester. Arret." -Type WARN
    Write-Log @log_params -Msg "=== Fin Test-SPOTargets ===" -Type INFO
    exit 0
}

# --- 2. Connexion initiale (1ere URL) pour reutiliser la session -----------
$dstsiteConnection = $null
$firstUrl          = $urls | Select-Object -First 1
Write-Log @log_params -Msg "Connexion initiale (Browser) sur : $firstUrl" -Type INFO
try {
    $dstsiteConnection = Connect-Site -Url $firstUrl -Browser -ErrorAction Stop
} catch {
    Write-Log @log_params -Msg "Echec de la connexion initiale sur '$firstUrl' : $($_.Exception.Message)" -Type ERROR
    throw "Echec de la connexion initiale sur '$firstUrl' : $($_.Exception.Message)"
}

# --- 3. Test de chaque URL -------------------------------------------------
$results = foreach ($url in $urls) {

    $sw = [System.Diagnostics.Stopwatch]::StartNew()

    # Validation syntaxique basique
    $parsed = $null
    if (-not [System.Uri]::TryCreate($url, [System.UriKind]::Absolute, [ref]$parsed) `
        -or $parsed.Scheme -notin @('http','https')) {
        Write-Log @log_params -Msg "[KO] $url  ->  URL invalide" -Type ERROR
        [pscustomobject]@{
            TargetSPOURL = $url
            Status       = 'InvalidUrl'
            SiteTitle    = ''
            Duration     = $sw.Elapsed
            Message      = "URL mal formee ou schema non http(s)"
        }
        continue
    }

    try {
        if ($url -eq $firstUrl) {
            $site = $dstsiteConnection
        } else {
            $site = Connect-Site -Url $url -UseCredentialsFrom $dstsiteConnection -ErrorAction Stop
        }

        $title = $null
        try { $title = $site.Title } catch { $title = '' }

        Write-Log @log_params -Msg ("[OK] {0}  ->  {1}" -f $url, $title) -Type INFO
        [pscustomobject]@{
            TargetSPOURL = $url
            Status       = 'Found'
            SiteTitle    = $title
            Duration     = $sw.Elapsed
            Message      = ''
        }
    }
    catch {
        $msg    = $_.Exception.Message
        $status = if ($msg -match '404|not\s*found|introuvable|does not exist') { 'NotFound' }
                  elseif ($msg -match '401|403|unauthorized|forbidden|access denied') { 'AccessDenied' }
                  else { 'Error' }

        Write-Log @log_params -Msg ("[KO] {0}  ->  {1} : {2}" -f $url, $status, $msg) -Type ERROR
        [pscustomobject]@{
            TargetSPOURL = $url
            Status       = $status
            SiteTitle    = ''
            Duration     = $sw.Elapsed
            Message      = $msg
        }
    }
    finally { $sw.Stop() }
}

# --- 4. Export du rapport CSV ----------------------------------------------
$reportFile = Join-Path $OutputDir ("{0}_TargetsCheck_{1}.csv" -f $batchBase, (Get-Date -Format 'yyyyMMdd_HHmmss'))
try {
    $results | Export-Csv -Path $reportFile -Delimiter $Delimiter -NoTypeInformation -Encoding UTF8
    Write-Log @log_params -Msg "Rapport CSV : $reportFile" -Type INFO
} catch {
    Write-Log @log_params -Msg "Echec de l'export CSV '$reportFile' : $($_.Exception.Message)" -Type ERROR
}

# --- 5. Rapport ------------------------------------------------------------
$ok = ($results | Where-Object Status -eq 'Found').Count
$ko = ($results | Where-Object Status -ne 'Found').Count
Write-Log @log_params -Msg "Resultat : $ok OK / $ko KO sur $($results.Count) URL(s)." -Type INFO
Write-Log @log_params -Msg "Log journalier : $LogFile" -Type INFO
Write-Log @log_params -Msg "=== Fin Test-SPOTargets ===" -Type INFO

if ($ko -gt 0) { exit 1 } else { exit 0 }
