<#
.SYNOPSIS
    Cree l'arborescence de dossiers cible (colonne TargetFolder) dans chaque site liste dans le CSV.

.DESCRIPTION
    Pour chaque ligne du CSV :
      - TargetFolder definit l'arborescence a creer sous la bibliotheque "Shared Documents"
        (ex. "General/Migration-Sharepoint-2026", "MonCanal/Sous-dossier", "Migration-2026", ...)
      - Chaque segment manquant est cree (creation incrementale, idempotente).
      - Un sous-dossier supplementaire est ensuite cree, nomme d'apres le dernier segment de SourcePath.

    Les lignes sont regroupees par site (TargetSPOURL) : une seule connexion PnP est ouverte
    par site distinct, puis reutilisee pour toutes les lignes de ce site.

    Les sorties sont isolees dans un sous-dossier dedie au script :
      - Logs   : C:\migrationFactory\logs\<NomDuScript>\
      - Rapport: C:\migrationFactory\output\<NomDuScript>\
    Le log est journalier (date du jour incluse dans le nom du fichier).

    Gestion des erreurs sur 2 niveaux :
      - Echec de connexion a un site  -> toutes les lignes de ce site sont marquees 'Failed', on passe au site suivant.
      - Echec de creation d'un dossier -> la ligne est marquee 'Failed', on passe a la ligne suivante.

.PARAMETER BatchFile
    Chemin complet vers le CSV, OU simple nom de fichier present dans le dossier 'input'
    de la migration factory (C:\migrationFactory\input).
    Colonnes attendues (separateur ';') : SourcePath, TargetType, TargetSPOURL, TargetFolder.

.EXAMPLE
    .\02-Create-SPOFolders.ps1 -BatchFile .\SPO-Input_Worker_1_20260625.csv

.EXAMPLE
    .\02-Create-SPOFolders.ps1 -BatchFile SPO-Input_Worker_1_20260625.csv
    # Le fichier est recherche automatiquement dans C:\migrationFactory\input
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$BatchFile
)

# Racine de la migration factory (en dur)
$FactoryRoot = 'C:\migrationFactory'

# Constantes internes
$LibraryRoot      = 'Shared Documents'
$ClientId         = '0f78653b-2b30-47f2-9d09-3c17709f118a'
$CredentialTarget = 'ShareGate3'

# --- 0. Pre-requis ---------------------------------------------------------
Import-Module PnP.PowerShell    -ErrorAction Stop
Import-Module CredentialManager -ErrorAction Stop

# Identifiant du script (ex: "02-Create-SPOFolders") => sert a nommer les sous-dossiers
$scriptTag = [System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Path)
if ([string]::IsNullOrWhiteSpace($scriptTag)) { $scriptTag = 'Create-SPOFolders' }

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

# Recuperation du credential depuis le Credential Manager
$cred = Get-StoredCredential -Target $CredentialTarget
if (-not $cred) {
    throw "Credential '$CredentialTarget' introuvable dans le Credential Manager. L'ajouter via cmdkey ou l'UI."
}

$Delimiter = ';'

# Log journalier : PrimagaZ_Create_Folder_M365_YYYYMMDD.log (date du jour)
$dayTag  = Get-Date -Format 'yyyyMMdd'
$execLog = Join-Path $LogsDir ("PrimagaZ_Create_Folder_M365_{0}.log" -f $dayTag)

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
$log_params = @{ Name = "New-SPOMigrationFolders" ; Logpath = $execLog }

Write-Log @log_params -Msg "=== Demarrage New-SPOMigrationFolders ===" -Type INFO
Write-Log @log_params -Msg "Batch       : $BatchFile"        -Type INFO
Write-Log @log_params -Msg "Bibliotheque: $LibraryRoot"      -Type INFO
Write-Log @log_params -Msg "ClientId    : $ClientId"         -Type INFO
Write-Log @log_params -Msg "Compte      : $($cred.UserName)" -Type INFO
Write-Log @log_params -Msg "Dossier log : $LogsDir"          -Type INFO
Write-Log @log_params -Msg "Dossier out : $OutputDir"        -Type INFO
Write-Log @log_params -Msg "Log         : $execLog"          -Type INFO

# --- 1. Lecture du CSV -----------------------------------------------------
$rows = Import-Csv -Path $BatchFile -Delimiter $Delimiter

foreach ($col in 'TargetFolder','TargetSPOURL','SourcePath') {
    if (-not ($rows | Get-Member -Name $col -MemberType NoteProperty)) {
        throw "Colonne '$col' absente de '$BatchFile'."
    }
}

# --- 2. Helper : creation incrementale d'une arborescence ------------------
function Ensure-PnPFolderPath {
    param(
        [Parameter(Mandatory)] $Connection,
        [Parameter(Mandatory)] [string]$RelSiteUrl,
        [Parameter(Mandatory)] [string]$LibraryRoot,
        [Parameter(Mandatory)] [string]$RelativePath
    )

    $segments = $RelativePath -split '[\\/]+' | Where-Object { $_ }
    $current  = $LibraryRoot
    $created  = $false

    foreach ($seg in $segments) {
        $childRel = "$RelSiteUrl/$current/$seg"
        $existing = Get-PnPFolder -Url $childRel -Connection $Connection -ErrorAction SilentlyContinue
        if (-not $existing) {
            $null = Add-PnPFolder -Name $seg -Folder $current -Connection $Connection -ErrorAction Stop
            $created = $true
        }
        $current = "$current/$seg"
    }

    [pscustomobject]@{ FullPath = $current ; Created = $created }
}

# --- 3. Helper : traitement d'une ligne (sur une connexion deja ouverte) ---
function Invoke-FolderRow {
    param(
        [Parameter(Mandatory)] $Connection,
        [Parameter(Mandatory)] $Row,
        [Parameter(Mandatory)] [string]$Url,
        [Parameter(Mandatory)] [hashtable]$LogParams,
        [Parameter(Mandatory)] [string]$LibraryRoot
    )

    $sw           = [System.Diagnostics.Stopwatch]::StartNew()
    $type         = "$($Row.TargetType)".Trim()
    $sourcePath   = "$($Row.SourcePath)".Trim()
    $targetFolder = "$($Row.TargetFolder)".Trim() -replace '\\','/' -replace '/+','/' -replace '^/+','' -replace '/+$',''

    # Sous-dossier additionnel : nom du dernier segment du SourcePath (ex. \\srv\share\Test -> "Test")
    $leafName = $null
    if ($sourcePath) {
        try { $leafName = Split-Path -Path $sourcePath -Leaf } catch { $leafName = $null }
    }

    if (-not $targetFolder) {
        Write-Log @LogParams -Msg "[WARN] $Url  ->  TargetFolder vide (ignore)" -Type WARN
        $sw.Stop()
        return [pscustomobject]@{ Url=$Url; Type=$type; TargetFolder=''; Leaf=$leafName; Status='Skipped'; Message='TargetFolder vide'; Duration=$sw.Elapsed }
    }

    try {
        $relSiteUrl = ([System.Uri]$Url).AbsolutePath.TrimEnd('/')

        # --- a) Arborescence TargetFolder (creation incrementale) ---
        $rootResult = Ensure-PnPFolderPath -Connection $Connection `
                                           -RelSiteUrl $relSiteUrl `
                                           -LibraryRoot $LibraryRoot `
                                           -RelativePath $targetFolder

        $rootStatus = if ($rootResult.Created) { 'Created' } else { 'Exists' }
        $rootLabel  = "$LibraryRoot/$targetFolder"
        if ($rootResult.Created) {
            Write-Log @LogParams -Msg "[OK] $Url  ->  $rootLabel cree" -Type INFO
        } else {
            Write-Log @LogParams -Msg "[=]  $Url  ->  $rootLabel (deja present)" -Type INFO
        }

        # --- b) Sous-dossier issu du SourcePath (leaf) ---
        $leafStatus  = 'Skipped'
        $leafMessage = ''
        if (-not $leafName) {
            $leafMessage = 'SourcePath vide ou non resoluble'
            Write-Log @LogParams -Msg "[WARN] $Url  ->  pas de leaf SourcePath, sous-dossier non cree" -Type WARN
        } else {
            $leafParent = "$LibraryRoot/$targetFolder"
            $leafRel    = "$relSiteUrl/$leafParent/$leafName"
            $existingLeaf = Get-PnPFolder -Url $leafRel -Connection $Connection -ErrorAction SilentlyContinue
            if ($existingLeaf) {
                Write-Log @LogParams -Msg "[=]  $Url  ->  $leafParent/$leafName (deja present)" -Type INFO
                $leafStatus = 'Exists'
            } else {
                $null = Add-PnPFolder -Name $leafName -Folder $leafParent -Connection $Connection -ErrorAction Stop
                Write-Log @LogParams -Msg "[OK] $Url  ->  $leafParent/$leafName cree" -Type INFO
                $leafStatus = 'Created'
            }
        }

        # Statut global agrege : Failed > Created > Exists > Skipped
        $status = if ($rootStatus -eq 'Created' -or $leafStatus -eq 'Created') { 'Created' }
                  elseif ($leafStatus -eq 'Skipped') { 'PartiallyCreated' }
                  else { 'Exists' }

        $sw.Stop()
        return [pscustomobject]@{
            Url          = $Url
            Type         = $type
            TargetFolder = $targetFolder
            Leaf         = $leafName
            Status       = $status
            Message      = $leafMessage
            Duration     = $sw.Elapsed
        }
    }
    catch {
        Write-Log @LogParams -Msg "[KO] $Url  ->  $($_.Exception.Message)" -Type ERROR
        $sw.Stop()
        return [pscustomobject]@{ Url=$Url; Type=$type; TargetFolder=$targetFolder; Leaf=$leafName; Status='Failed'; Message=$_.Exception.Message; Duration=$sw.Elapsed }
    }
}

# --- 4. Boucle de creation (regroupement par site) -------------------------
$results = [System.Collections.Generic.List[object]]::new()

# Regroupe les lignes par TargetSPOURL : 1 connexion par site distinct
$groups = $rows | Group-Object -Property { "$($_.TargetSPOURL)".Trim() }

foreach ($group in $groups) {

    $url      = $group.Name
    $siteRows = $group.Group

    # Lignes sans URL : ignorees, mais tracees dans le rapport
    if (-not $url) {
        foreach ($row in $siteRows) {
            $type = "$($row.TargetType)".Trim()
            $tf   = "$($row.TargetFolder)".Trim()
            $sp   = "$($row.SourcePath)".Trim()
            $leaf = if ($sp) { try { Split-Path -Path $sp -Leaf } catch { $null } } else { $null }
            $results.Add([pscustomobject]@{ Url=''; Type=$type; TargetFolder=$tf; Leaf=$leaf; Status='Skipped'; Message='TargetSPOURL vide'; Duration=[TimeSpan]::Zero })
        }
        continue
    }

    Write-Log @log_params -Msg "--- Site : $url  ($($siteRows.Count) ligne(s)) ---" -Type INFO

    # --- Connexion au site (1 fois pour toutes ses lignes) ---
    $conn = $null
    try {
        $conn = Connect-PnPOnline -Url $url `
                                  -ClientId $ClientId `
                                  -Credentials $cred `
                                  -ReturnConnection `
                                  -ErrorAction Stop
    }
    catch {
        # Echec connexion site -> toutes les lignes du site en Failed, on passe au site suivant
        $connMsg = $_.Exception.Message
        Write-Log @log_params -Msg "[KO] $url  ->  Echec de connexion au site : $connMsg" -Type ERROR
        foreach ($row in $siteRows) {
            $type = "$($row.TargetType)".Trim()
            $tf   = "$($row.TargetFolder)".Trim() -replace '\\','/' -replace '/+','/' -replace '^/+','' -replace '/+$',''
            $sp   = "$($row.SourcePath)".Trim()
            $leaf = if ($sp) { try { Split-Path -Path $sp -Leaf } catch { $null } } else { $null }
            $results.Add([pscustomobject]@{ Url=$url; Type=$type; TargetFolder=$tf; Leaf=$leaf; Status='Failed'; Message="Connexion site KO : $connMsg"; Duration=[TimeSpan]::Zero })
        }
        continue
    }

    # --- Traitement de chaque ligne du site (connexion reutilisee) ---
    try {
        foreach ($row in $siteRows) {
            $res = Invoke-FolderRow -Connection $conn `
                                    -Row $row `
                                    -Url $url `
                                    -LogParams $log_params `
                                    -LibraryRoot $LibraryRoot
            $results.Add($res)
        }
    }
    finally {
        if ($conn) {
            try { Disconnect-PnPOnline -Connection $conn -ErrorAction SilentlyContinue } catch {}
        }
    }
}

# --- 5. Export du rapport CSV ----------------------------------------------
$reportFile = Join-Path $OutputDir ("{0}_FoldersCreate_{1}.csv" -f $batchBase, (Get-Date -Format 'yyyyMMdd_HHmmss'))
try {
    $results | Export-Csv -Path $reportFile -Delimiter $Delimiter -NoTypeInformation -Encoding UTF8
    Write-Log @log_params -Msg "Rapport CSV : $reportFile" -Type INFO
} catch {
    Write-Log @log_params -Msg "Echec de l'export CSV '$reportFile' : $($_.Exception.Message)" -Type ERROR
}

# --- 6. Rapport ------------------------------------------------------------
$created = ($results | Where-Object Status -in 'Created','PartiallyCreated').Count
$exists  = ($results | Where-Object Status -eq 'Exists').Count
$failed  = ($results | Where-Object Status -eq 'Failed').Count
$skipped = ($results | Where-Object Status -eq 'Skipped').Count

Write-Log @log_params -Msg "Bilan : $created cree(s) / $exists deja present(s) / $failed echec(s) / $skipped ignore(s)" -Type INFO
Write-Log @log_params -Msg "Log journalier : $execLog" -Type INFO
Write-Log @log_params -Msg "=== Fin New-SPOMigrationFolders ===" -Type INFO

if ($failed -gt 0) { exit 1 } else { exit 0 }
