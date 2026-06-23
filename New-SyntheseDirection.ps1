<#
.SYNOPSIS
    Genere un espace de synthese Excel consolide (direction) a partir des CSV
    d'inventaire produits par Get-PathTooLong_Mig_Sync_V4.ps1.
.DESCRIPTION
    Parcourt chaque CSV 'Inventaire_*.csv' (1 CSV = 1 ligne du mapping = 1 file share),
    calcule la volumetrie reelle et la synthese des chemins trop longs (> 400 car.,
    soit Migrable = False), puis produit un classeur Excel a 2 onglets :
      - Synthese globale  : 1 ligne par file share + ligne TOTAL (KPI volumetrie)
      - Chemins trop longs : detail consolide des elements non migrables
    La colonne 'File Share' contient le chemin UNC complet du partage (racine reelle),
    reconstruite de facon deterministe via 'CheminSource' moins 'ProfondeurSource' segments.
    Perimetre "migrable" = tout SAUF les extensions bloquees (les chemins > 400 car.
    restent comptes comme migrables car ils seront remedies avant migration).
    Objectif : communiquer a la direction une vue claire du volume a migrer et
    faciliter le suivi des remediations avant migration.
.PARAMETER ReportsPath
    Dossier racine contenant les sous-dossiers de file shares et leurs CSV.
    Defaut : .\path_too_long_Reports
.PARAMETER OutputXlsx
    Chemin du classeur Excel de synthese.
    Defaut : <ReportsPath>\Synthese_Direction_<timestamp>.xlsx
.PARAMETER SeuilMigration
    Seuil de chemin de migration (informatif, affiche dans l'en-tete). Defaut : 400
.NOTES
    Projet : PrimaGAZ - Migration FileShare vers M365 | Phase : 01 - Assessment
#>
param(
    [ValidateNotNullOrEmpty()][string]$ReportsPath = (Join-Path $PSScriptRoot "path_too_long_Reports"),
    [string]$OutputXlsx,
    [ValidateRange(1,32767)][int]$SeuilMigration = 400
)

# ======================================================================
# PREREQUIS
# ======================================================================

if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "[ERREUR] Module 'ImportExcel' requis." -ForegroundColor Red
    Write-Host "  Install-Module ImportExcel -Scope CurrentUser" -ForegroundColor Cyan
    throw "Module ImportExcel manquant."
}
Import-Module ImportExcel -ErrorAction Stop

if (-not (Test-Path -Path $ReportsPath)) { throw "Dossier introuvable : '$ReportsPath'" }

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
if ([string]::IsNullOrWhiteSpace($OutputXlsx)) {
    $OutputXlsx = Join-Path $ReportsPath "Synthese_Direction_$timestamp.xlsx"
}
$outputDir = Split-Path -Parent $OutputXlsx
if ($outputDir -and -not (Test-Path $outputDir)) {
    New-Item -Path $outputDir -ItemType Directory -Force | Out-Null
}
if (Test-Path $OutputXlsx) { Remove-Item $OutputXlsx -Force }

# ======================================================================
# HELPERS
# ======================================================================

function Write-Log {
    param([string]$Message,[ValidateSet("INFO","WARN","ERROR","SUCCESS")][string]$Level="INFO")
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $prefix = switch ($Level) { "INFO"{"[INFO]"} "WARN"{"[WARN]"} "ERROR"{"[ERROR]"} "SUCCESS"{"[OK]"} }
    $line = "$ts $prefix $Message"
    switch ($Level) {
        "ERROR"   { Write-Host $line -ForegroundColor Red }
        "WARN"    { Write-Host $line -ForegroundColor Yellow }
        "SUCCESS" { Write-Host $line -ForegroundColor Green }
        default   { Write-Host $line }
    }
}

# Conversion Mo -> Go (TailleMo est en Mo dans le CSV source ; vide pour les dossiers)
function ConvertTo-Go { param([double]$Mo) return [math]::Round($Mo / 1024, 2) }

# Somme robuste de TailleMo (ignore vides/dossiers, gere virgule ou point decimal)
function Get-SumTailleMo {
    param([object[]]$Items)
    $sum = 0.0
    foreach ($it in $Items) {
        $v = $it.TailleMo
        if ($null -ne $v -and "$v" -ne '') {
            $d = 0.0
            $normalized = ("$v" -replace ',', '.')
            if ([double]::TryParse($normalized, [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$d)) {
                $sum += $d
            }
        }
    }
    return $sum
}

# Booleen robuste : "True"/"true"/$true -> $true
function ConvertTo-Bool {
    param($Value)
    if ($Value -is [bool]) { return $Value }
    return ("$Value".Trim() -ieq 'True')
}

# Racine UNC du file share, reconstruite de facon deterministe :
#   racine = CheminSource auquel on retire 'ProfondeurSource' segments.
# La ligne racine n'est pas presente dans le CSV (prof min observee = 1),
# donc on remonte depuis l'element de profondeur la plus FAIBLE (le plus sur).
# Cela evite de capter un fichier court (ex. 'Thumbs.db') comme racine.
# Garde-fou minimal : ne jamais remonter au-dessus de \\serveur\premier-segment.
function Get-CheminFileShare {
    param([object[]]$Items,[string]$Fallback)

    $ref = $null; $refProf = [int]::MaxValue
    foreach ($it in $Items) {
        $cs = "$($it.CheminSource)"
        if ([string]::IsNullOrWhiteSpace($cs)) { continue }
        $prof = 0
        if (-not [int]::TryParse("$($it.ProfondeurSource)", [ref]$prof)) { continue }
        if ($prof -lt $refProf) { $refProf = $prof; $ref = $cs.TrimEnd('\') }
    }
    if ([string]::IsNullOrWhiteSpace($ref)) { return $Fallback }

    # Plancher de securite = \\serveur\premier-segment (jamais en dessous)
    $plancher = $ref
    if ($ref.StartsWith('\\')) {
        $segs = $ref.Substring(2) -split '\\' | Where-Object { $_ -ne '' }
        if ($segs.Count -ge 2) { $plancher = '\\' + $segs[0] + '\' + $segs[1] }
    }

    # Remonter de refProf segments (chaque segment = un niveau sous la racine)
    $racine = $ref
    for ($i = 0; $i -lt $refProf; $i++) {
        if ($racine.Length -le $plancher.Length) { break }
        $parent = Split-Path -Parent $racine
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent.Length -lt $plancher.Length) { break }
        $racine = $parent
    }
    return $racine
}

# ======================================================================
# COULEURS (alignees sur Convert-InventaireCsvToExcel.ps1)
# ======================================================================

$vertFond    = [System.Drawing.Color]::FromArgb(198, 239, 206)
$vertTexte   = [System.Drawing.Color]::FromArgb(0, 97, 0)
$rougeFond   = [System.Drawing.Color]::FromArgb(255, 199, 206)
$rougeTexte  = [System.Drawing.Color]::FromArgb(156, 0, 6)
$orangeFond  = [System.Drawing.Color]::FromArgb(255, 235, 156)
$orangeTexte = [System.Drawing.Color]::FromArgb(156, 101, 0)
$bleuEntete  = [System.Drawing.Color]::FromArgb(68, 114, 196)
$grisTotal   = [System.Drawing.Color]::FromArgb(217, 217, 217)
$blancTexte  = [System.Drawing.Color]::White

# Nom de colonne dynamique (le seuil est parametrable) - calcule UNE fois
$colTropLongs = "Nb chemins trop longs (>$SeuilMigration)"

Write-Host ""
Write-Host "========================================================" -ForegroundColor Cyan
Write-Host " Espace de synthese direction - Volumetrie & chemins longs" -ForegroundColor Cyan
Write-Host "========================================================" -ForegroundColor Cyan
Write-Host "  Reports : $ReportsPath" -ForegroundColor White
Write-Host "  Sortie  : $OutputXlsx" -ForegroundColor White
Write-Host "  Seuil   : $SeuilMigration car. (chemin de migration SPO)" -ForegroundColor White
Write-Host ""

# ======================================================================
# DECOUVERTE DES CSV (1 CSV le plus recent par file share)
# ======================================================================

Write-Log "Recherche des CSV 'Inventaire_*.csv' dans : $ReportsPath" "INFO"

$allCsv = @(Get-ChildItem -Path $ReportsPath -Recurse -Filter "Inventaire_*.csv" -File -ErrorAction SilentlyContinue)
if ($allCsv.Count -eq 0) {
    throw "Aucun fichier 'Inventaire_*.csv' trouve sous '$ReportsPath'."
}
Write-Log "Fichiers trouves : $($allCsv.Count)" "INFO"

# Un file share = un sous-dossier ; on garde le CSV le plus recent par dossier parent
$csvParGroupe = $allCsv |
    Group-Object { $_.DirectoryName } |
    ForEach-Object { $_.Group | Sort-Object LastWriteTime -Descending | Select-Object -First 1 }

$csvRetenus = @($csvParGroupe | Sort-Object FullName)
Write-Log "CSV retenus (le plus recent par dossier) : $($csvRetenus.Count)" "SUCCESS"

# ======================================================================
# ANALYSE PAR CSV
# ======================================================================

$syntheseGlobale  = New-Object 'System.Collections.Generic.List[object]'
$cheminsTropLongs = New-Object 'System.Collections.Generic.List[object]'

# Compteurs TOTAL
$totElements = 0; $totFichiers = 0; $totDossiers = 0
$totElementsMigrables = 0
$totTailleMo = 0.0; $totTailleMigrableMo = 0.0
$totTropLongs = 0; $totTailleTropLongsMo = 0.0

foreach ($csv in $csvRetenus) {

    Write-Log "Analyse : $($csv.FullName)" "INFO"

    $rows = @(Import-Csv -Path $csv.FullName -Delimiter ';' -Encoding UTF8)
    if ($rows.Count -eq 0) {
        Write-Log "  CSV vide, ignore." "WARN"
        continue
    }

    # Nom du file share : depuis la colonne, sinon depuis le dossier parent
    $nomFS = ($rows | Where-Object { $_.NomFileShare } | Select-Object -First 1).NomFileShare
    if ([string]::IsNullOrWhiteSpace($nomFS)) { $nomFS = Split-Path -Leaf $csv.DirectoryName }

    # Chemin UNC complet (racine reelle) : reconstruit via CheminSource - ProfondeurSource
    $cheminFS = Get-CheminFileShare -Items $rows -Fallback $nomFS

    $typeCible = ($rows | Where-Object { $_.TargetType } | Select-Object -First 1).TargetType
    if ([string]::IsNullOrWhiteSpace($typeCible)) { $typeCible = "(inconnu)" }

    $fichiers   = @($rows | Where-Object { $_.TypeElement -eq 'Fichier' })
    $dossiers   = @($rows | Where-Object { $_.TypeElement -eq 'Dossier' })

    # Volumetrie
    # Perimetre migrable = tout SAUF extensions bloquees (les chemins > 400 restent inclus,
    # car ils seront remedies avant migration). Le nombre et la taille partagent ce perimetre.
    $tailleMo          = Get-SumTailleMo -Items $rows
    $migrablesExtOk    = @($rows | Where-Object { -not (ConvertTo-Bool $_.ExtensionBloquee) })
    $nbMigrables       = $migrablesExtOk.Count
    $tailleMigrableMo  = Get-SumTailleMo -Items $migrablesExtOk

    # Chemins trop longs = Migrable = False (strictement > 400 car. migration)
    $tropLongs         = @($rows | Where-Object { -not (ConvertTo-Bool $_.Migrable) })
    $tailleTropLongsMo = Get-SumTailleMo -Items $tropLongs

    # Ligne de synthese - File Share = UNC complet ; Nb elements migrables avant Taille migrable
    $ligne = [ordered]@{
        'File Share'            = $cheminFS
        'Type cible'            = $typeCible
        'Nb elements'           = $rows.Count
        'Nb fichiers'           = $fichiers.Count
        'Nb dossiers'           = $dossiers.Count
        'Taille totale (Go)'    = ConvertTo-Go $tailleMo
        'Nb elements migrables' = $nbMigrables
        'Taille migrable (Go)'  = ConvertTo-Go $tailleMigrableMo
    }
    $ligne[$colTropLongs] = $tropLongs.Count
    $ligne['Taille concernee par chemins longs (Go)'] = ConvertTo-Go $tailleTropLongsMo
    $syntheseGlobale.Add([PSCustomObject]$ligne)

    # Detail consolide des chemins trop longs (pour le suivi remediation)
    foreach ($it in ($tropLongs | Sort-Object { [int]$_.DepassementMigration } -Descending)) {
        $cheminsTropLongs.Add([PSCustomObject][ordered]@{
            'File Share'             = $cheminFS
            'Type cible'             = $typeCible
            'Type'                   = $it.TypeElement
            'Nom'                    = $it.Nom
            'Extension'              = $it.Extension
            'Taille (Mo)'            = if ("$($it.TailleMo)" -ne '') { $it.TailleMo } else { "" }
            'Longueur chemin SPO'    = [int]$it.LongueurCheminMigration
            'Depassement (car.)'     = [int]$it.DepassementMigration
            'Profondeur'             = [int]$it.ProfondeurSource
            'Chemin sur le serveur'  = $it.CheminSource
        })
    }

    # Cumuls TOTAL
    $totElements          += $rows.Count
    $totFichiers          += $fichiers.Count
    $totDossiers          += $dossiers.Count
    $totElementsMigrables += $nbMigrables
    $totTailleMo          += $tailleMo
    $totTailleMigrableMo  += $tailleMigrableMo
    $totTropLongs         += $tropLongs.Count
    $totTailleTropLongsMo += $tailleTropLongsMo

    Write-Log ("  -> {0} elements ({1} migrables) | {2} Go total | {3} chemins trop longs ({4} Go)" -f `
        $rows.Count, $nbMigrables, (ConvertTo-Go $tailleMo), $tropLongs.Count, (ConvertTo-Go $tailleTropLongsMo)) "SUCCESS"
}

if ($syntheseGlobale.Count -eq 0) {
    throw "Aucune donnee exploitable dans les CSV retenus."
}

# Ligne TOTAL
$ligneTotal = [ordered]@{
    'File Share'            = 'TOTAL'
    'Type cible'            = ''
    'Nb elements'           = $totElements
    'Nb fichiers'           = $totFichiers
    'Nb dossiers'           = $totDossiers
    'Taille totale (Go)'    = ConvertTo-Go $totTailleMo
    'Nb elements migrables' = $totElementsMigrables
    'Taille migrable (Go)'  = ConvertTo-Go $totTailleMigrableMo
}
$ligneTotal[$colTropLongs] = $totTropLongs
$ligneTotal['Taille concernee par chemins longs (Go)'] = ConvertTo-Go $totTailleTropLongsMo
$syntheseGlobale.Add([PSCustomObject]$ligneTotal)

# ======================================================================
# ONGLET 1 - SYNTHESE GLOBALE
# ======================================================================

Write-Host ""
Write-Log "Creation onglet : Synthese globale..." "INFO"

$titreSynth = "SYNTHESE MIGRATION FILESHARE -> M365  (seuil chemin SPO : $SeuilMigration car.)"

$xlPkg = $syntheseGlobale | Export-Excel -Path $OutputXlsx `
    -WorksheetName "Synthese globale" `
    -TableName "TblSynthese" -TableStyle Medium2 -AutoFilter `
    -Title $titreSynth -TitleBold -TitleSize 14 -PassThru

$ws = $xlPkg.Workbook.Worksheets["Synthese globale"]
$maxRow = $ws.Dimension.End.Row
$maxCol = $ws.Dimension.End.Column

# Ligne d'en-tete de tableau (apres le titre)
$headerRow = $ws.Dimension.Start.Row + 1
for ($col = 1; $col -le $maxCol; $col++) {
    $cell = $ws.Cells[$headerRow, $col]
    $cell.Style.Font.Bold = $true
    $cell.Style.Font.Color.SetColor($blancTexte)
    $cell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
    $cell.Style.Fill.BackgroundColor.SetColor($bleuEntete)
}

# Mise en valeur de la ligne TOTAL (derniere ligne)
for ($col = 1; $col -le $maxCol; $col++) {
    $cell = $ws.Cells[$maxRow, $col]
    $cell.Style.Font.Bold = $true
    $cell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
    $cell.Style.Fill.BackgroundColor.SetColor($grisTotal)
}

# Coloration de la colonne "Nb chemins trop longs" : rouge si > 0
$colIdxTropLongs = -1
for ($col = 1; $col -le $maxCol; $col++) {
    if ($ws.Cells[$headerRow, $col].Text -like 'Nb chemins trop longs*') { $colIdxTropLongs = $col; break }
}
if ($colIdxTropLongs -gt 0) {
    for ($row = $headerRow + 1; $row -le $maxRow; $row++) {
        $cell = $ws.Cells[$row, $colIdxTropLongs]
        $valNum = 0
        if ([int]::TryParse($cell.Text, [ref]$valNum)) {
            if ($valNum -gt 0) {
                $cell.Style.Font.Bold = $true
                $cell.Style.Font.Color.SetColor($rougeTexte)
                if ($row -ne $maxRow) {
                    $cell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                    $cell.Style.Fill.BackgroundColor.SetColor($rougeFond)
                }
            } elseif ($row -ne $maxRow) {
                $cell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                $cell.Style.Fill.BackgroundColor.SetColor($vertFond)
                $cell.Style.Font.Color.SetColor($vertTexte)
            }
        }
    }
}

# Format millier pour les colonnes de taille (Go)
for ($col = 1; $col -le $maxCol; $col++) {
    $h = $ws.Cells[$headerRow, $col].Text
    if ($h -like '*(Go)*' ) {
        for ($row = $headerRow + 1; $row -le $maxRow; $row++) {
            $ws.Cells[$row, $col].Style.Numberformat.Format = '#,##0.00'
        }
    }
}

$ws.View.FreezePanes($headerRow + 1, 1)
for ($col = 1; $col -le $maxCol; $col++) {
    $ws.Column($col).AutoFit()
    if ($ws.Column($col).Width -gt 80) { $ws.Column($col).Width = 80 }
}
Close-ExcelPackage $xlPkg
Write-Log "  -> $($syntheseGlobale.Count - 1) file share(s) + ligne TOTAL" "SUCCESS"

# ======================================================================
# ONGLET 2 - CHEMINS TROP LONGS
# ======================================================================

Write-Log "Creation onglet : Chemins trop longs..." "INFO"

if ($cheminsTropLongs.Count -gt 0) {
    $titreCTL = "CHEMINS TROP LONGS A REMEDIER AVANT MIGRATION (> $SeuilMigration car.) - $($cheminsTropLongs.Count) element(s)"

    $xlPkg = $cheminsTropLongs | Export-Excel -Path $OutputXlsx `
        -WorksheetName "Chemins trop longs" `
        -TableName "TblCheminsLongs" -TableStyle Medium3 -AutoFilter `
        -Title $titreCTL -TitleBold -TitleSize 14 -PassThru

    $ws = $xlPkg.Workbook.Worksheets["Chemins trop longs"]
    $maxRow = $ws.Dimension.End.Row
    $maxCol = $ws.Dimension.End.Column
    $headerRow = $ws.Dimension.Start.Row + 1

    for ($col = 1; $col -le $maxCol; $col++) {
        $cell = $ws.Cells[$headerRow, $col]
        $cell.Style.Font.Bold = $true
        $cell.Style.Font.Color.SetColor($blancTexte)
        $cell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
        $cell.Style.Fill.BackgroundColor.SetColor($bleuEntete)
    }

    # Colonne Depassement en rouge gras
    $colDep = -1
    for ($col = 1; $col -le $maxCol; $col++) {
        if ($ws.Cells[$headerRow, $col].Text -like 'Depassement*') { $colDep = $col; break }
    }
    if ($colDep -gt 0) {
        for ($row = $headerRow + 1; $row -le $maxRow; $row++) {
            $cell = $ws.Cells[$row, $colDep]
            $valNum = 0
            if ([int]::TryParse($cell.Text, [ref]$valNum) -and $valNum -gt 0) {
                $cell.Style.Font.Bold = $true
                $cell.Style.Font.Color.SetColor($rougeTexte)
            }
        }
    }

    $ws.View.FreezePanes($headerRow + 1, 1)
    for ($col = 1; $col -le $maxCol; $col++) {
        $ws.Column($col).AutoFit()
        if ($ws.Column($col).Width -gt 70) { $ws.Column($col).Width = 70 }
    }
    Close-ExcelPackage $xlPkg
    Write-Log "  -> $($cheminsTropLongs.Count) element(s) a remedier" "WARN"
} else {
    $msgOK = @([PSCustomObject]@{ Message = "Aucun chemin trop long (> $SeuilMigration car.). Tous les elements sont migrables." })
    $xlPkg = $msgOK | Export-Excel -Path $OutputXlsx `
        -WorksheetName "Chemins trop longs" -AutoSize -PassThru
    $ws = $xlPkg.Workbook.Worksheets["Chemins trop longs"]
    $ws.Cells[2, 1].Style.Font.Size = 14
    $ws.Cells[2, 1].Style.Font.Color.SetColor($vertTexte)
    $ws.Row(1).Hidden = $true
    $ws.Column(1).Width = 80
    Close-ExcelPackage $xlPkg
    Write-Log "  -> 0 chemin trop long" "SUCCESS"
}

# ======================================================================
# REORGANISATION DES ONGLETS
# ======================================================================

$xlPkg = Open-ExcelPackage -Path $OutputXlsx
$worksheets = $xlPkg.Workbook.Worksheets
if ($null -ne $worksheets["Synthese globale"]) { $worksheets.MoveToStart("Synthese globale") }
$worksheets["Synthese globale"].View.TabSelected = $true
foreach ($w in $worksheets) {
    if ($w.Name -ne "Synthese globale") { $w.View.TabSelected = $false }
}
Close-ExcelPackage $xlPkg

# ======================================================================
# TERMINE
# ======================================================================

$fileSize = [math]::Round((Get-Item $OutputXlsx).Length / 1KB, 0)
Write-Host ""
Write-Host "========================================================" -ForegroundColor Green
Write-Host " Fichier de synthese genere avec succes" -ForegroundColor Green
Write-Host "========================================================" -ForegroundColor Green
Write-Host "  Fichier : $OutputXlsx ($fileSize Ko)" -ForegroundColor White
Write-Host ""
Write-Host "  KPI globaux :" -ForegroundColor Cyan
Write-Host ("    File shares analyses        : {0}" -f ($syntheseGlobale.Count - 1)) -ForegroundColor White
Write-Host ("    Elements totaux             : {0}" -f $totElements) -ForegroundColor White
Write-Host ("    Elements migrables (h. ext) : {0}" -f $totElementsMigrables) -ForegroundColor White
Write-Host ("    Taille totale               : {0} Go" -f (ConvertTo-Go $totTailleMo)) -ForegroundColor White
Write-Host ("    Taille migrable (hors ext.) : {0} Go" -f (ConvertTo-Go $totTailleMigrableMo)) -ForegroundColor White
Write-Host ("    Chemins trop longs (>{0})   : {1}" -f $SeuilMigration, $totTropLongs) `
    -ForegroundColor $(if ($totTropLongs -gt 0) { "Red" } else { "Green" })
Write-Host ("    Taille concernee (longs)    : {0} Go" -f (ConvertTo-Go $totTailleTropLongsMo)) -ForegroundColor White
Write-Host ""
Write-Host "  Onglets :" -ForegroundColor Cyan
Write-Host "    1. Synthese globale   - 1 ligne par file share (UNC complet) + TOTAL" -ForegroundColor White
Write-Host "    2. Chemins trop longs - detail consolide a remedier" -ForegroundColor White
Write-Host ""
