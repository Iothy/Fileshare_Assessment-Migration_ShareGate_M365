<#
.SYNOPSIS
    Dashboard décisionnel multi-FileShare pour l'assessment PrimaGAZ.

.DESCRIPTION
    Ce script lit les fichiers CSV générés par les autres scripts d'assessment
    et produit un dashboard HTML exécutif présentable pour le client PrimaGAZ.

    Sans -FileShareMapping : génère 1 rapport global unique.
    Avec -FileShareMapping : génère 1 rapport par ligne du mapping (1 par FileShare) +
        1 rapport global confidentiel + 1 index DSI.
    Avec -FileShareMapping -SplitByLevel1 : mode legacy, rapports par dossier N1.

.PARAMETER CheminOutput
    Chemin du dossier Output contenant les fichiers CSV (par défaut : .\Output).

.PARAMETER FileShareMapping
    Chemin du fichier FileShareMapping.csv (optionnel).
    Si fourni, enrichit les rapports avec le nom du FileShare, le type d'usage,
    la cible M365, l'owner et son email.
    Format : CheminUNC;NomFileShare;TypeUsage;CibleM365;Owner;EmailOwner;Description

.PARAMETER ReportOutputPath
    Dossier de destination des rapports HTML et CSV annexes
    (par défaut : sous-dossier _Reports dans CheminOutput).

.PARAMETER SplitByLevel1
    Si activé, génère des rapports séparés par FileShare et dossier de niveau 1.

.EXAMPLE
    .\Export-AssessmentReport.ps1

.EXAMPLE
    .\Export-AssessmentReport.ps1 -CheminOutput "C:\Resultats"

.EXAMPLE
    .\Export-AssessmentReport.ps1 -CheminOutput ".\Output" -FileShareMapping ".\Config\FileShareMapping.csv"

    Genere 1 rapport par ligne du FileShareMapping (+1 GLOBAL +1 INDEX).
    Fichiers attendus dans Output/ : Inventaire_FileShare_<NomFS>_*.csv, sous-dossiers Permissions/ et AgeFichiers/.
    Rapports HTML/CSV annexes écrits par défaut dans Output/_Reports/.

.EXAMPLE
    .\Export-AssessmentReport.ps1 -CheminOutput ".\Output" -FileShareMapping ".\Config\FileShareMapping.csv" -SplitByLevel1

    Mode legacy : rapports par dossier N1 au sein de chaque FileShare.

.EXAMPLE
    .\Export-AssessmentReport.ps1 -CheminOutput ".\Output" -FileShareMapping ".\Config\FileShareMapping.csv" -ReportOutputPath ".\Reports\2026-05-12"

.NOTES
    Projet  : PrimaGAZ - Migration FileShare vers M365
    Phase   : 01 - Assessment
    Version : v3.4
    Auteur  : Equipe Migration Avanade

    CHANGELOG
    ---------
    v3.4 (2026-05-12)
      - Nouveau paramètre optionnel -ReportOutputPath pour définir le dossier de sortie
        des rapports HTML et CSV annexes (InheritanceHotspots, IdentitiesInventory).
      - Sans -Run et sans -ReportOutputPath, les rapports sont maintenant écrits dans
        <CheminOutput>\_Reports (creation automatique, fallback sur CheminOutput en cas d'échec).
      - Permissions NTFS : filtrage ACL par racine UNC normalisée (StartsWith insensible à la casse),
        log du nombre de lignes matchées par FileShare, section toujours affichee si le CSV
        permissions est disponible (message explicite quand aucune ligne ne correspond).

    v3.3 (2026-05-10)
      - Correction 1 : Pyramide des ages fichiers.
        Nouvelle section "Pyramide des ages des fichiers" dans chaque rapport par FileShare (mode mapping).
        Source : Inventaire_FileShare_Detail_<NomFS>_*.csv (colonne DateModification).
        Buckets : <1 an, 1-2 ans, 2-5 ans, 5-10 ans, >10 ans.
        Decision client : seuls les fichiers <2 ans seront migres — mise en evidence via stat-boxes vertes/grises.
        Fonction dediee : Build-SectionPyramideAges.
      - Correction 2 : Section "Extensions a auditer" masquee.
        Le lien nav, la barre chart, la ligne synthese et la ligne plan d'action ne sont plus affiches.
        La lecture du CSV et la variable $dataExtensions sont conservees pour ne rien casser.
      - Correction 3 : Detection groupes AD elargie.
        Nouveaux patterns : SHVE\USR_BTQ_*, Domain Admins/Users/Computers/Guests,
        Enterprise/Schema Admins, _G_/_U_ entre underscores (nomenclature PrimaGAZ).
      - Correction 4 : Bug "N1 impactes toujours = 1" dans les hotspots de permissions.
        Nouveau parametre $CheminUNCRoot passe a Build-SectionPermissions et Build-FullReport.
        Le N1 est calcule en retirant la racine UNC du share avant le split, revelant les vrais
        dossiers N1 (ex : "F8 CONSEILS JURIDIQUES" au lieu de "DRH_2").
      - Correction 5 (patch) : Profondeur relative a la racine logique du FileShare.
        Helpers Get-DepthRelativeToShareRoot et Get-N1RelativeToShareRoot ajoutes.
        La profondeur dans la table hotspots et la stat-box "Profondeur max rupture" est
        desormais relative a la racine du FileShare (racine = 1) quand $CheminUNCRoot est fourni,
        au lieu d'etre calculee sur le chemin UNC complet (qui donnait 3 pour la racine).
        Le calcul N1 est migre de l'inline vers Get-N1RelativeToShareRoot.
        Mode SplitByLevel1 : $pfx/$fsRoot2 propages comme CheminUNCRoot dans Build-FullReport.

    v3.2 (2026-05-05)
      - Nettoyage perimetre client :
        1. Owner supprime : plus aucune mention de Owner / A_DEFINIR / a.definir@ dans le HTML.
           La lecture Owner/EmailOwner dans FileShareMapping.csv est conservee (utilisee pour TypeUsage/CibleM365).
        2. Duree migration estimee supprimee : stat-box et fonction Format-DurationReadable retires.
           Le bloc Readiness conserve exactement 5 stat-boxes utiles.
        3. Caracteres invalides supprimes : gere nativement par ShareGate.
           Chargement de CaracteresInvalides_*.csv commente, li synthese, entree plan d'action et lien nav retires.
        4. Fichiers volumineux supprimes : Get-LargeFiles.ps1 non execute par le client.
           Chargement de FichiersVolumineux_*.csv commente, li synthese, graphique, entree plan et lien nav retires.
        5. Blocklist extensions : l'export consomme uniquement le CSV produit par Get-BlockedExtensions.ps1
           (pas de refiltrage interne), conforme a la liste officielle PrimaGAZ.
      - Permissions NTFS enrichies :
        6. Nouveau sous-bloc « Hotspots de ruptures d'heritage » dans #section-permissions :
           stat-boxes (nb dossiers casses, N1 impactes, % arborescence, profondeur max), chart Top 10 N1,
           table Top 50 hotspots, CSV InheritanceHotspots_<NomFS>_<Timestamp>.csv.
        7. Nouveau sous-bloc « Inventaire des identites AD » dans #section-permissions :
           stat-boxes (total identites, groupes, utilisateurs), table complete (sans limite),
           CSV IdentitiesInventory_<NomFS>_<Timestamp>.csv.
        8. Readiness global : ruptures d'heritage comptees comme Warnings.
        9. Plan d'action : ligne P1 « Arbitrer les ruptures d'heritage NTFS » si N > 0.
       10. Chart « Top problemes detectes » : barre « Heritage casse » (orange/warn).
       11. Footer : v3.2 — Permissions NTFS : hotspots heritage + inventaire identites AD.

    v3.1 (2026-05-05)
      - Bug 1 : Correction du « Migration Readiness % » qui affichait le % d'erreurs au lieu du % de readiness.
      - Bug 2 : Filtrage strict du NomFileShare dans Get-FileShareCsvSet (_newest utilise désormais une regex
                pour éviter que CPRM corresponde à CpRM_Signatures et DATALAKE à DATALAKE_RCT).
      - Bug 3 : Build-SectionInventaire affiche une ligne [Racine] quand le share n'a pas de sous-dossiers N1
                (ex. CALCULATEUR : 8 fichiers à la racine).
      - Bug 4 : Format adaptatif KB/MB/GB/TB pour les volumes et la durée estimée (0 GB → 29 KB).
      - Bug 5 : Le graphique Top 10 dossiers N1 par volume inclut désormais les dossiers à 0 octet (en bas).
      - Bug 6 : Le sous-titre Owner utilise le format (email) au lieu de <email> pour éviter le double-encodage.
      - Bug 7 : Build-SectionAgeFichiers affiche « Aucun fichier ancien détecté » quand il n'y a pas de données.
      - Bug 8 : Ajout d'une alerte visuelle rouge en haut du rapport quand Owner = A_DEFINIR.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$CheminOutput = (Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Definition) "Output"),

    [Parameter(Mandatory = $false, HelpMessage = "Chemin du fichier FileShareMapping.csv")]
    [string]$FileShareMapping = "",

    [Parameter(Mandatory = $false, HelpMessage = "Dossier de destination des rapports HTML et CSV annexes (par défaut : sous-dossier _Reports/ dans CheminOutput)")]
    [string]$ReportOutputPath = "",

    [Parameter(Mandatory = $false, HelpMessage = "Générer des rapports séparés par dossier de niveau 1")]
    [switch]$SplitByLevel1,

    [Parameter(Mandatory = $false, HelpMessage = "Objet Run issu de New-AssessmentRun (structure hiérarchique v3.0)")]
    [PSCustomObject]$Run = $null
)

# ============================================================
# INITIALISATION
# ============================================================

Import-Module (Join-Path -Path $PSScriptRoot -ChildPath "Modules\PrimaGAZ.Assessment.psm1") -Force
$outputModulePath = Join-Path -Path $PSScriptRoot -ChildPath "Modules\PrimaGAZ.Output.psm1"
if (Test-Path $outputModulePath) { Import-Module $outputModulePath -Force }

# Routing : si $Run fourni, utiliser son dossier csv comme source et _Reports/ pour les rapports HTML
if ($null -ne $Run) {
    $CheminOutput = $Run.Csv
    # BaseOutput = 2 niveaux au-dessus du run folder (BaseOutput\Scope\Timestamp)
    $baseOutputDir = Split-Path (Split-Path $Run.Path -Parent) -Parent
    $reportOutputPath = Join-Path $baseOutputDir "_Reports"
    if (-not (Test-Path $reportOutputPath)) {
        try { New-Item -ItemType Directory -Path $reportOutputPath -Force | Out-Null }
        catch { Write-Warning "Impossible de créer le dossier _Reports : $_"; $reportOutputPath = $Run.Path }
    }
} else {
    if (-not [string]::IsNullOrWhiteSpace($ReportOutputPath)) {
        if ([System.IO.Path]::IsPathRooted($ReportOutputPath)) {
            $reportOutputPath = $ReportOutputPath
        } else {
            $reportOutputPath = Join-Path (Get-Location).Path $ReportOutputPath
        }
        if (-not (Test-Path $reportOutputPath)) {
            try { New-Item -ItemType Directory -Path $reportOutputPath -Force | Out-Null }
            catch { Write-Warning "Impossible de créer le dossier de rapports '$reportOutputPath' : $_"; $reportOutputPath = $CheminOutput }
        }
    } else {
        $reportOutputPath = Join-Path $CheminOutput "_Reports"
        if (-not (Test-Path $reportOutputPath)) {
            try { New-Item -ItemType Directory -Path $reportOutputPath -Force | Out-Null }
            catch { Write-Warning "Impossible de créer le dossier _Reports '$reportOutputPath' : $_"; $reportOutputPath = $CheminOutput }
        }
    }
}

$timestamp    = Get-Date -Format "yyyyMMdd_HHmmss"
$MAX_ROWS     = 50
$VOL_ELEVE_TB = 5

if (-not (Test-Path $CheminOutput)) {
    Write-Error "Le dossier '$CheminOutput' n'existe pas. Exécutez d'abord les scripts d'analyse."
    exit 1
}

Write-Log "=== Génération du dashboard décisionnel PrimaGAZ v3.4 ===" "INFO"

# ============================================================
# SOUS-DOSSIERS SPECIAUX
# ============================================================
$permOutputPath = Join-Path $CheminOutput "Permissions"
$ageOutputPath  = Join-Path $CheminOutput "AgeFichiers"

# ============================================================
# LECTURE DES CSV
# ============================================================

function Import-CsvSafe {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path $Path)) { return @() }
    try { return @(Import-Csv -Path $Path -Delimiter ";" -Encoding UTF8) }
    catch { Write-Log "Erreur lecture CSV : $Path — $_" "WARN"; return @() }
}

# Sélectionner le CSV le plus récent correspondant à un pattern (sans regex stricte)
function Select-NewestCsv {
    param([string]$Folder, [string]$Filter, [string]$FallbackName = "")
    if (-not (Test-Path $Folder)) { return $null }
    $found = Get-ChildItem -Path $Folder -Filter $Filter -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($found) { return $found }
    if ($FallbackName) {
        $fb = Join-Path $Folder $FallbackName
        return Get-Item $fb -ErrorAction SilentlyContinue
    }
    return $null
}

$csvInventaire   = Select-NewestCsv $CheminOutput "Inventaire_FileShare_*.csv"      "Inventaire_FileShare.csv"
# Prefer consolidated (no NomFS segment) if present, else take any
if ($csvInventaire -and $csvInventaire.Name -match '^Inventaire_FileShare_\d{8}') {
    # consolidated found, keep it
} elseif ($csvInventaire -and $csvInventaire.Name -match '^Inventaire_FileShare_.+_\d{8}') {
    # only per-FS files exist, grab the most recent any
    $csvInventaire = Get-ChildItem -Path $CheminOutput -Filter "Inventaire_FileShare_*.csv" |
        Sort-Object LastWriteTime -Descending | Select-Object -First 1
}

$csvVolumineux   = $null  # v3.2 — FichiersVolumineux supprime (Get-LargeFiles.ps1 non execute par le client)
# $csvVolumineux   = Select-NewestCsv $CheminOutput "FichiersVolumineux_*.csv"        "FichiersVolumineux.csv"
$csvInvalides    = $null  # v3.2 — CaracteresInvalides supprime (gere nativement par ShareGate)
# $csvInvalides    = Select-NewestCsv $CheminOutput "CaracteresInvalides_*.csv"       "CaracteresInvalides.csv"
# Support both spellings: CheminsTropLongs and ChemainsTropLongs
$csvChemins      = Select-NewestCsv $CheminOutput "CheminsTropLongs_*.csv"          "CheminsTropLongs.csv"
if (-not $csvChemins) { $csvChemins = Select-NewestCsv $CheminOutput "ChemainsTropLongs_*.csv" "ChemainsTropLongs.csv" }
$csvExtensions   = Select-NewestCsv $CheminOutput "ExtensionsBloquees_*.csv"        "ExtensionsBloquees.csv"
$csvDoublons     = Select-NewestCsv $CheminOutput "FichiersDupliques_*.csv"         "FichiersDupliques.csv"
# Permissions : look in Permissions/ subfolder first
$csvPermissions  = Select-NewestCsv $permOutputPath "Permissions_NTFS_*.csv"       "Permissions_NTFS.csv"
if (-not $csvPermissions) { $csvPermissions = Select-NewestCsv $CheminOutput "Permissions_NTFS_*.csv" "Permissions_NTFS.csv" }
$csvSidResolution = Select-NewestCsv $CheminOutput "SID_Resolution_*.csv"          "SID_Resolution.csv"
if (-not $csvSidResolution) { $csvSidResolution = Select-NewestCsv $permOutputPath "SID_Resolution_*.csv" "SID_Resolution.csv" }
$csvAccessDenied  = Select-NewestCsv $permOutputPath "AccessDenied_*.csv"          "AccessDenied.csv"
if (-not $csvAccessDenied)  { $csvAccessDenied  = Select-NewestCsv $CheminOutput  "AccessDenied_*.csv" "AccessDenied.csv" }
$csvHomeOwnership = Select-NewestCsv $CheminOutput "HomeDirectoryOwnership_*.csv"  "HomeDirectoryOwnership.csv"
$csvHomeStats     = Select-NewestCsv $CheminOutput "HomeDirectoryStats_*.csv"      "HomeDirectoryStats.csv"

$dataInventaire    = if ($csvInventaire)    { Import-CsvSafe $csvInventaire.FullName }    else { @() }
$dataVolumineux    = if ($csvVolumineux)    { Import-CsvSafe $csvVolumineux.FullName }    else { @() }
$dataInvalides     = if ($csvInvalides)     { Import-CsvSafe $csvInvalides.FullName }     else { @() }
$dataChemins       = if ($csvChemins)       { Import-CsvSafe $csvChemins.FullName }       else { @() }
$dataDoublons      = if ($csvDoublons)      { Import-CsvSafe $csvDoublons.FullName }      else { @() }
$dataPermissions   = if ($csvPermissions)   { Import-CsvSafe $csvPermissions.FullName }   else { @() }
$dataExtensions    = if ($csvExtensions)    { Import-CsvSafe $csvExtensions.FullName }    else { @() }
$dataSidResolution = if ($csvSidResolution) { Import-CsvSafe $csvSidResolution.FullName } else { @() }
$dataAccessDenied  = if ($csvAccessDenied)  { Import-CsvSafe $csvAccessDenied.FullName }  else { @() }
$dataHomeOwnership = if ($csvHomeOwnership) { Import-CsvSafe $csvHomeOwnership.FullName } else { @() }
$dataHomeStats     = if ($csvHomeStats)     { Import-CsvSafe $csvHomeStats.FullName }     else { @() }

# Filtrer les lignes "aucun résultat"
$dataVolumineux    = @($dataVolumineux    | Where-Object { $_.ResultatAnalyse -notmatch "^Aucun" })
$dataInvalides     = @($dataInvalides     | Where-Object { $_.ResultatAnalyse -notmatch "^Aucun" -and $_.ProblemeType -ne "DoublePoint" }) # DoublePoint : règle supprimée (aucune restriction MS documentée pour '..' dans SPO)
$dataChemins       = @($dataChemins       | Where-Object { $_.ResultatAnalyse -notmatch "^Aucun" })
$dataDoublons      = @($dataDoublons      | Where-Object { $_.ResultatAnalyse -notmatch "^Aucun" })
$dataExtensions    = @($dataExtensions    | Where-Object { $_.ResultatAnalyse -notmatch "^Aucun" })
$dataHomeOwnership = @($dataHomeOwnership | Where-Object { $_.ResultatAnalyse -notmatch "^Aucun" })
$dataHomeStats     = @($dataHomeStats     | Where-Object { $_.ResultatAnalyse -notmatch "^Aucun" })

# ============================================================
# LECTURE DU FILESHARE MAPPING
# ============================================================

$fsMapping = @()
if (-not [string]::IsNullOrWhiteSpace($FileShareMapping) -and (Test-Path $FileShareMapping)) {
    $fsMapping = @(Import-CsvSafe $FileShareMapping)
    Write-Log "FileShareMapping chargé : $($fsMapping.Count) entrées" "INFO"
} else {
    Write-Log "Aucun FileShareMapping fourni — mode sans enrichissement FileShare" "WARN"
}

# ============================================================
# FONCTIONS UTILITAIRES
# ============================================================

function HtmlEnc {
    param([string]$s)
    if ([string]::IsNullOrEmpty($s)) { return "" }
    return $s.Replace("&","&amp;").Replace("<","&lt;").Replace(">","&gt;").Replace('"',"&quot;").Replace("'","&#39;")
}

function ConvertTo-FileUrl {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return "" }
    return "file:///" + $Path.Replace("\", "/")
}

function Format-TruncatedPath {
    param([string]$Path, [int]$MaxLen = 80)
    if ([string]::IsNullOrWhiteSpace($Path)) { return "" }
    if ($MaxLen -lt 4) { return $Path.Substring(0, [Math]::Min($MaxLen, $Path.Length)) }
    if ($Path.Length -le $MaxLen) { return $Path }
    return "..." + $Path.Substring($Path.Length - ($MaxLen - 3))
}

function Get-ParentFolder {
    param([string]$FullPath)
    if ([string]::IsNullOrWhiteSpace($FullPath)) { return "" }
    try {
        $parent = [System.IO.Path]::GetDirectoryName($FullPath)
        if ([string]::IsNullOrWhiteSpace($parent)) { return $FullPath }
        return $parent
    } catch { return $FullPath }
}

function Get-EmplacementCell {
    param([string]$CheminComplet)
    if ([string]::IsNullOrWhiteSpace($CheminComplet)) { return "<td class='path-cell'>—</td>" }
    $parent  = Get-ParentFolder $CheminComplet
    if ([string]::IsNullOrWhiteSpace($parent)) { $parent = $CheminComplet }
    $url     = ConvertTo-FileUrl $parent
    $display = Format-TruncatedPath $parent 80
    $tip     = HtmlEnc $parent
    $disp    = HtmlEnc $display
    return "<td class='path-cell'><a href='$url' title='$tip' target='_blank'>$disp</a></td>"
}

function Get-ReadinessInfo {
    param([long]$BlockingErrors, [long]$TotalItems)
    if ($TotalItems -le 0) {
        return @{ Level="INCONNU"; Color="#95a5a6"; BgColor="#ecf0f1"; Icon="&#x2753;"; Text="DONNEES INSUFFISANTES"; Pct=0; ErrPct=0 }
    }
    $errPct = [math]::Round(100.0 * $BlockingErrors / $TotalItems, 1)
    $readinessPct = [math]::Max(0, [math]::Round(100.0 - $errPct, 1))
    if ($errPct -le 1) {
        return @{ Level="PRET";        Color="#27ae60"; BgColor="#eafaf1"; Icon="&#x1F7E2;"; Text="PRET A MIGRER";          Pct=$readinessPct; ErrPct=$errPct }
    } elseif ($errPct -le 5) {
        return @{ Level="REMEDIATION"; Color="#e67e22"; BgColor="#fef9e7"; Icon="&#x1F7E1;"; Text="REMEDIATION NECESSAIRE"; Pct=$readinessPct; ErrPct=$errPct }
    } else {
        return @{ Level="CRITIQUE";    Color="#e74c3c"; BgColor="#fdedec"; Icon="&#x1F534;"; Text="BLOCKERS CRITIQUES";     Pct=$readinessPct; ErrPct=$errPct }
    }
}

function Get-FsForPath {
    param([string]$Path, [array]$Mapping)
    if (-not $Path -or $Mapping.Count -eq 0) { return $null }
    $best = $null; $bestLen = 0
    foreach ($m in $Mapping) {
        if (-not $m.CheminUNC) { continue }
        $root = $m.CheminUNC.TrimEnd('\')
        if ($Path -like "$root*" -and $root.Length -gt $bestLen) {
            $best = $m; $bestLen = $root.Length
        }
    }
    return $best
}

function Get-N1Folder {
    param([string]$Path, [string]$FsRoot)
    if ([string]::IsNullOrEmpty($Path) -or [string]::IsNullOrEmpty($FsRoot)) { return "" }
    $root = $FsRoot.TrimEnd('\')
    if (-not $Path.StartsWith($root, [System.StringComparison]::OrdinalIgnoreCase)) { return "" }
    $rel = $Path.Substring($root.Length).TrimStart('\', '/')
    if ([string]::IsNullOrWhiteSpace($rel)) { return "" }
    $parts = $rel -split '[/\\]', 2
    return $parts[0]
}

function Normalize-PathForCompare {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return "" }
    $p = $Path.Trim().Replace('/', '\')
    $p = $p.TrimEnd('\')
    return $p.ToLowerInvariant()
}

function Get-PermissionsForScope {
    param(
        [array]$DataPermissions,
        [string]$ScopeRoot = "",
        [string]$NomFS = ""
    )
    if ($DataPermissions.Count -eq 0) { return @() }

    $pathCol = if ($DataPermissions[0].PSObject.Properties.Name -contains "CheminDossier") { "CheminDossier" } else { "Path" }
    $scopeNorm = Normalize-PathForCompare $ScopeRoot
    $matched = [System.Collections.Generic.List[object]]::new()

    if (-not [string]::IsNullOrWhiteSpace($scopeNorm)) {
        foreach ($row in $DataPermissions) {
            $rawPath = $row.$pathCol
            $pathNorm = Normalize-PathForCompare $rawPath
            if ([string]::IsNullOrWhiteSpace($pathNorm)) { continue }
            if ($pathNorm -eq $scopeNorm -or $pathNorm.StartsWith($scopeNorm + "\")) {
                $matched.Add($row)
            }
        }
    }

    if ($matched.Count -eq 0 -and -not [string]::IsNullOrWhiteSpace($NomFS)) {
        $byFs = @($DataPermissions | Where-Object { $_.PSObject.Properties['NomFileShare'] -and $_.NomFileShare -eq $NomFS })
        if ($byFs.Count -gt 0) { return $byFs }
    }

    return @($matched)
}

# v3.3 patch — Profondeur relative a la racine logique du FileShare
# ShareRoot = CheminUNC de la ligne mapping (ex: \\server\EQUIPE\Juridique).
# Racine du share = profondeur 1 ; chaque segment supplementaire = +1.
# Si ShareRoot est vide ou si le chemin n'est pas sous ShareRoot, on retombe sur la
# profondeur UNC absolue (retrocompatibilite mode standard / rapport global).
function Get-DepthRelativeToShareRoot {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$ShareRoot
    )
    # Fallback: absolute UNC depth (used when ShareRoot is blank or path is outside the share)
    $uncDepth = { ($Path.TrimStart('\', '/') -split '[/\\]' | Where-Object { $_ }).Count }
    if ([string]::IsNullOrWhiteSpace($Path) -or [string]::IsNullOrWhiteSpace($ShareRoot)) { return & $uncDepth }
    # Normalize to lowercase backslash-separated for comparison only; display uses original casing
    $rootNorm = $ShareRoot.TrimEnd('\', '/').Replace('/', '\').ToLowerInvariant()
    $pathNorm = $Path.TrimEnd('\', '/').Replace('/', '\').ToLowerInvariant()
    if ($pathNorm -eq $rootNorm) { return 1 }
    if (-not $pathNorm.StartsWith($rootNorm + '\')) {
        # chemin hors share — fallback : profondeur depuis la racine UNC
        return & $uncDepth
    }
    $relative = $Path.Substring($ShareRoot.Length).Trim('\', '/')
    if ([string]::IsNullOrWhiteSpace($relative)) { return 1 }
    $segments = @($relative -split '[/\\]' | Where-Object { $_ })
    return 1 + $segments.Count
}

# v3.3 patch — Dossier N1 relatif a la racine logique du FileShare
# Retourne le premier segment sous ShareRoot, ou "[Racine]" si le chemin EST la racine,
# ou "" si le chemin est hors share (fallback).
function Get-N1RelativeToShareRoot {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$ShareRoot
    )
    if ([string]::IsNullOrWhiteSpace($Path) -or [string]::IsNullOrWhiteSpace($ShareRoot)) { return "" }
    # Normalize to lowercase backslash-separated for comparison only; display uses original casing
    $rootNorm = $ShareRoot.TrimEnd('\', '/').Replace('/', '\').ToLowerInvariant()
    $pathNorm = $Path.TrimEnd('\', '/').Replace('/', '\').ToLowerInvariant()
    if ($pathNorm -eq $rootNorm) { return "[Racine]" }
    if (-not $pathNorm.StartsWith($rootNorm + '\')) { return "" }
    $relative = $Path.Substring($ShareRoot.Length).Trim('\', '/')
    if ([string]::IsNullOrWhiteSpace($relative)) { return "[Racine]" }
    $segments = @($relative -split '[/\\]' | Where-Object { $_ })
    if ($segments.Count -eq 0) { return "[Racine]" }
    return $segments[0]
}

function ConvertTo-LongSafe {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return 0 }
    $v = $Value.Trim() -replace '\s', '' -replace ',', '.'
    try {
        $d = [double]::Parse($v, [System.Globalization.CultureInfo]::InvariantCulture)
        if ($d -lt 0) { return 0 }
        return [long]$d
    } catch { return 0 }
}

function Format-VolumeReadable {
    param([double]$VolumeGB)
    if ($VolumeGB -ge 1024)           { return "$([math]::Round($VolumeGB/1024, 2)) TB" }
    if ($VolumeGB -ge 1)              { return "$([math]::Round($VolumeGB, 1)) GB" }
    if ($VolumeGB -ge (1.0 / 1024))   { return "$([math]::Round($VolumeGB * 1024, 0)) MB" }
    return "$([math]::Round($VolumeGB * 1048576, 0)) KB"
}

# Format-DurationReadable supprimee en v3.2 (stat-box "Duree migration estimee" retiree)

function Get-FileShareCsvSet {
    param(
        [string]$OutputRoot,
        [string]$NomFS
    )
    $permRoot = Join-Path $OutputRoot "Permissions"
    $ageRoot  = Join-Path $OutputRoot "AgeFichiers"

    # Helper: pick newest CSV matching a filter in a folder.
    # $ExactNomFS: when provided, enforces that the NomFS segment in the filename is followed
    # immediately by a timestamp (_\d{8}_\d{6}), preventing e.g. CPRM from matching CpRM_Signatures.
    function _newest { param([string]$Dir,[string]$Filter,[string]$ExactNomFS="")
        if (-not (Test-Path $Dir)) { return $null }
        $candidates = Get-ChildItem -Path $Dir -Filter $Filter -ErrorAction SilentlyContinue
        if ($ExactNomFS) {
            $escaped = [regex]::Escape($ExactNomFS)
            $candidates = $candidates | Where-Object { $_.Name -match "(?i)_${escaped}_\d{8}_\d{6}\." }
        }
        $candidates | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    }

    $result = [PSCustomObject]@{
        Inventaire       = (_newest $OutputRoot "Inventaire_FileShare_${NomFS}_*.csv"         $NomFS)
        InventaireDetail = (_newest $OutputRoot "Inventaire_FileShare_Detail_${NomFS}_*.csv"  $NomFS)
        Volumineux       = $null  # v3.2 — FichiersVolumineux supprime (Get-LargeFiles.ps1 non execute par le client)
        # Volumineux    = (_newest $OutputRoot "FichiersVolumineux_${NomFS}_*.csv"           $NomFS)
        Chemins          = (_newest $OutputRoot "CheminsTropLongs_${NomFS}_*.csv"             $NomFS)
        Extensions       = (_newest $OutputRoot "ExtensionsBloquees_${NomFS}_*.csv"           $NomFS)
        Invalides        = $null  # v3.2 — CaracteresInvalides supprime (gere nativement par ShareGate)
        # Invalides     = (_newest $OutputRoot "CaracteresInvalides_${NomFS}_*.csv"          $NomFS)
        Doublons         = (_newest $OutputRoot "FichiersDupliques_${NomFS}_*.csv"            $NomFS)
        Permissions      = (_newest $permRoot   "Permissions_NTFS_${NomFS}_*.csv"            $NomFS)
        AccessDenied     = (_newest $permRoot   "AccessDenied_${NomFS}_*.csv"                $NomFS)
        SidResolution    = (_newest $OutputRoot "SID_Resolution_${NomFS}_*.csv"              $NomFS)
        HomeOwnership    = (_newest $OutputRoot "HomeDirectoryOwnership_${NomFS}_*.csv"      $NomFS)
        HomeStats        = (_newest $OutputRoot "HomeDirectoryStats_${NomFS}_*.csv"          $NomFS)
        AgeSummary       = $null
        AgeDetails       = @()
        DossiersLVT      = (_newest $OutputRoot "Dossiers_LVT_${NomFS}_*.csv"               $NomFS)
    }

    # Support the "ChemainsTropLongs" typo fallback
    if (-not $result.Chemins) {
        $result.Chemins = (_newest $OutputRoot "ChemainsTropLongs_${NomFS}_*.csv" $NomFS)
    }

    # AgeSummary: AgeFichiers/_SUMMARY_<NomFS>.csv
    if (Test-Path $ageRoot) {
        $sumPath = Join-Path $ageRoot "_SUMMARY_${NomFS}.csv"
        if (Test-Path $sumPath) { $result.AgeSummary = Get-Item $sumPath }

        # AgeDetails: all NonModifie_*.csv under AgeFichiers/<NomFS>/
        $fsDirAge = Join-Path $ageRoot $NomFS
        if (Test-Path $fsDirAge) {
            $result.AgeDetails = @(Get-ChildItem -Path $fsDirAge -Recurse -Filter "NonModifie_*.csv" -ErrorAction SilentlyContinue)
        }
    }

    return $result
}

function Load-FsData {
    param([object]$CsvFile, [array]$GlobalData, [string]$NomFS, [string]$CheminUNC)
    if ($CsvFile -and (Test-Path $CsvFile.FullName)) {
        $d = @(Import-CsvSafe $CsvFile.FullName)
        if ($d.Count -gt 0) { return $d }
    }
    # fallback: filter global data by NomFileShare column, or by path prefix
    if ($GlobalData.Count -eq 0) { return @() }
    $byFs = @($GlobalData | Where-Object { $_.PSObject.Properties['NomFileShare'] -and $_.NomFileShare -eq $NomFS })
    if ($byFs.Count -gt 0) { return $byFs }
    if ($CheminUNC) {
        $pfx = $CheminUNC.TrimEnd('\')
        return @($GlobalData | Where-Object { $_.CheminComplet -like "$pfx*" -or $_.CheminAnalyse -like "$pfx*" })
    }
    return @()
}

# ============================================================
# CSS PARTAGE
# ============================================================

function Get-CommonCss {
    return @"
* { box-sizing: border-box; margin: 0; padding: 0; }
body { font-family: 'Segoe UI', Tahoma, sans-serif; background: #f5f7fa; color: #333; font-size: 14px; }
a { color: #2980b9; text-decoration: none; }
a:hover { text-decoration: underline; }
#nav-sticky { position: sticky; top: 0; z-index: 100; background: #1e3a5f; padding: 8px 20px; display: flex; flex-wrap: wrap; gap: 4px; box-shadow: 0 2px 6px rgba(0,0,0,0.3); }
#nav-sticky a { color: #d0e8ff; font-size: 12px; padding: 4px 10px; border-radius: 12px; white-space: nowrap; transition: background 0.2s; }
#nav-sticky a:hover { background: rgba(255,255,255,0.2); text-decoration: none; }
.page-header { background: linear-gradient(135deg, #1e3a5f, #2980b9); color: white; padding: 30px 40px; margin-bottom: 24px; }
.page-header h1 { font-size: 26px; font-weight: 700; margin-bottom: 6px; }
.page-header p { opacity: 0.88; font-size: 14px; }
.section { background: white; border-radius: 8px; padding: 24px 28px; margin: 0 20px 20px; box-shadow: 0 2px 6px rgba(0,0,0,0.08); }
.section h2 { color: #1e3a5f; border-bottom: 2px solid #2980b9; padding-bottom: 10px; margin-bottom: 16px; font-size: 18px; }
.section h3 { color: #2c3e50; margin: 16px 0 10px; font-size: 15px; }
.stat-row { display: flex; flex-wrap: wrap; gap: 12px; margin-bottom: 16px; }
.stat-box { background: #eaf4fb; border-left: 4px solid #2980b9; padding: 14px 20px; border-radius: 6px; min-width: 130px; }
.stat-box.danger { background: #fdedec; border-color: #e74c3c; }
.stat-box.warn   { background: #fef9e7; border-color: #e67e22; }
.stat-box.ok     { background: #eafaf1; border-color: #27ae60; }
.stat-box .number { font-size: 22px; font-weight: 700; color: #1e3a5f; }
.stat-box.danger .number { color: #e74c3c; }
.stat-box.warn   .number { color: #e67e22; }
.stat-box.ok     .number { color: #27ae60; }
.stat-box .label { font-size: 11px; color: #666; text-transform: uppercase; margin-top: 4px; }
table { width: 100%; border-collapse: collapse; margin-top: 12px; font-size: 13px; }
th { background: #1e3a5f; color: white; padding: 9px 10px; text-align: left; font-weight: 600; }
td { padding: 7px 10px; border-bottom: 1px solid #eee; vertical-align: top; }
tr:nth-child(even) td { background: #f8fbfd; }
tr:hover td { background: #e8f4fb !important; }
.table-footer td { background: #f0f7ff !important; font-style: italic; color: #555; font-size: 12px; }
.tbl-total td { font-weight: 700 !important; background: #e8f4fd !important; }
.section-counter { font-size: 12px; color: #666; margin-bottom: 8px; }
.badge { display: inline-block; padding: 2px 8px; border-radius: 12px; font-size: 11px; font-weight: 600; }
.badge-error { background: #fdecea; color: #c0392b; border: 1px solid #e74c3c; }
.badge-warn  { background: #fef5e7; color: #d35400; border: 1px solid #e67e22; }
.badge-ok    { background: #eafaf1; color: #1e8449; border: 1px solid #27ae60; }
.badge-info  { background: #eaf4fb; color: #1a5276; border: 1px solid #2980b9; }
.path-cell { max-width: 300px; word-break: break-all; font-size: 12px; }
.path-cell a { color: #2980b9; }
.readiness-block { border-radius: 10px; padding: 24px 32px; display: flex; align-items: center; gap: 30px; margin-bottom: 20px; }
.readiness-icon { font-size: 56px; line-height: 1; }
.readiness-pct { font-size: 36px; font-weight: 800; }
.chart-stacked { display: flex; height: 32px; border-radius: 6px; overflow: hidden; margin: 12px 0; width: 100%; }
.chart-stacked .seg { display: flex; align-items: center; justify-content: center; font-size: 11px; font-weight: 600; color: white; overflow: hidden; min-width: 0; }
.seg-error { background: #e74c3c; }
.seg-warn  { background: #e67e22; }
.seg-ok    { background: #27ae60; }
.chart-horiz { margin: 8px 0; }
.horiz-row { display: flex; align-items: center; gap: 8px; margin-bottom: 6px; }
.horiz-label { width: 220px; font-size: 12px; color: #444; text-align: right; flex-shrink: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.horiz-track { flex: 1; background: #ecf0f1; border-radius: 4px; height: 18px; overflow: hidden; }
.horiz-fill { height: 100%; border-radius: 4px; }
.horiz-count { width: 50px; font-size: 12px; color: #555; }
.plan-table th:first-child { width: 60px; }
.priority-p0 { color: #c0392b; font-weight: 700; }
.priority-p1 { color: #d35400; font-weight: 700; }
.priority-p2 { color: #d4ac0d; font-weight: 700; }
.priority-p3 { color: #27ae60; font-weight: 700; }
.badge-volume { background: #fef9e7; color: #d35400; border: 1px solid #e67e22; font-size: 10px; padding: 1px 6px; border-radius: 10px; margin-left: 6px; }
.report-footer { text-align: center; color: #999; margin: 30px 20px 20px; font-size: 12px; padding: 16px; background: white; border-radius: 8px; }
.unmapped-alert { background: #fdedec; border: 1px solid #e74c3c; border-radius: 6px; padding: 12px 16px; margin: 8px 0; }
@media print {
  @page { size: A4 landscape; margin: 1cm; }
  body { background: white; font-size: 11px; }
  #nav-sticky { display: none; }
  .section { margin: 0 0 12px 0; box-shadow: none; border: 1px solid #ddd; break-inside: avoid; }
  .page-header { background: #1e3a5f !important; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
  tr { break-inside: avoid; }
}
@media (max-width: 768px) {
  .stat-row { gap: 8px; }
  .stat-box { min-width: 100px; }
  .horiz-label { width: 120px; }
  .readiness-block { flex-direction: column; text-align: center; }
}
"@
}

# ============================================================
# CONSTRUCTEURS DE SECTIONS HTML
# ============================================================

function Build-SectionReadiness {
    param(
        [long]$BlockingErrors,
        [long]$Warnings,
        [long]$TotalItems,
        [double]$VolumeGB,
        [string]$Scope = "Global"
    )
    $okCount = [math]::Max(0, $TotalItems - $BlockingErrors - $Warnings)
    $volStr = Format-VolumeReadable $VolumeGB

    return @"
<div class="section" id="section-readiness">
  <h2>&#x1F3AF; Migration Readiness</h2>
  <div class="stat-row">
    <div class="stat-box"><div class="number">$TotalItems</div><div class="label">Items analyses</div></div>
    <div class="stat-box danger"><div class="number">$BlockingErrors</div><div class="label">Erreurs bloquantes</div></div>
    <div class="stat-box warn"><div class="number">$Warnings</div><div class="label">Warnings</div></div>
    <div class="stat-box ok"><div class="number">$okCount</div><div class="label">Items OK</div></div>
    <div class="stat-box"><div class="number">$volStr</div><div class="label">Volume total</div></div>
  </div>
</div>
"@
}

function Build-SectionSummary {
    param(
        [long]$TotalFichiers,
        [double]$VolumeGB,
        [long]$BlockingErrors,
        [long]$Warnings,
        [long]$NbExtensions,
        [long]$NbInvalides,
        [long]$NbChemins,
        [long]$NbVolumineux,
        [long]$NbDoublons,
        [long]$NbRupturesHeritage = 0,
        [string]$Scope = "Global",
        [string]$NomFS = ""
    )
    $volStr        = Format-VolumeReadable $VolumeGB
    $nomFsDisplay  = if ([string]::IsNullOrWhiteSpace($NomFS)) { "PrimaGAZ" } else { $NomFS }
    $errorLabel    = if ($BlockingErrors -eq 1) { "erreur bloquante" } else { "erreurs bloquantes" }
    $warningLabel  = if ($Warnings -eq 1) { "point de vigilance" } else { "points de vigilance" }
    $cheminLabel   = if ($NbChemins -eq 1) { "chemin trop long" } else { "chemins trop longs" }
    $doublonLabel  = if ($NbDoublons -eq 1) { "doublon" } else { "doublons" }
    $ruptureLabel  = if ($NbRupturesHeritage -eq 1) { "dossier avec héritage cassé" } else { "dossiers avec héritage cassé" }
    $ruptureVerb   = if ($NbRupturesHeritage -eq 1) { "nécessite" } else { "nécessitent" }
    $txtR = if ($BlockingErrors -eq 0) {
        "Le périmètre est techniquement prêt à migrer."
    } elseif ($BlockingErrors -le 50) {
        "Une phase de remédiation ciblée est nécessaire avant le lancement de la migration."
    } else {
        "Une phase de remédiation significative est requise avant migration ; un planning de traitement par lot est recommandé."
    }

    $liRuptures = if ($NbRupturesHeritage -gt 0) { "<li>&#x1F510; <strong>$NbRupturesHeritage $ruptureLabel</strong> &mdash; $ruptureVerb un arbitrage avant migration (site dédié / bibliothèque dédiée / consolidation arbo)</li>" } else { "" }

    return @"
<div class="section" id="section-summary">
  <h2>&#x1F4DD; Synthèse Executive</h2>
  <p style="line-height:1.7;font-size:14px;">
    Le FileShare <strong>$(HtmlEnc $nomFsDisplay)</strong> représente <strong>$TotalFichiers fichiers</strong> pour un volume total de <strong>$volStr</strong>.<br>
    L'audit a identifié <strong style="color:#e74c3c;">$BlockingErrors $errorLabel</strong> à remédier avant migration et <strong style="color:#e67e22;">$Warnings $warningLabel</strong> à traiter.<br>
    $txtR
  </p>
  <ul style="margin:16px 0 0 20px;line-height:2.2;">
    <li>&#x1F4CF; <strong>$NbChemins $cheminLabel</strong> &mdash; restructuration arborescente requise</li>
    <li>&#x1F4D1; <strong>$NbDoublons $doublonLabel</strong> &mdash; déduplication recommandée avant migration</li>
    $liRuptures
  </ul>
</div>
"@
}

function Build-SectionCharts {
    param(
        [array]$DataInvalides,
        [array]$DataExtensions,
        [array]$DataChemins,
        [array]$DataVolumineux,
        [array]$DataInventaire,
        [long]$NbRupturesHeritage = 0
    )
    # --- Top 10 types de problemes ---
    $problemCounts = @{}
    foreach ($r in $DataInvalides)  {
        $k = $r.ProblemeType
        if ($k) { $problemCounts[$k] = (($problemCounts[$k] -as [int]) + 1) }
    }
    # v3.3 — Extensions bloquees ne sont plus affichees dans le chart (suppression affichage)
    if ($DataChemins.Count      -gt 0) { $problemCounts["Chemins trop longs"]  = $DataChemins.Count }
    if ($NbRupturesHeritage     -gt 0) { $problemCounts["Heritage casse"]      = $NbRupturesHeritage }

    $top10 = $problemCounts.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 10
    $maxP  = if ($top10) { ($top10 | Measure-Object Value -Maximum).Maximum } else { 1 }
    if ($maxP -le 0) { $maxP = 1 }

    $chartProblems = ""
    foreach ($entry in $top10) {
        $pct = [math]::Round(100.0 * $entry.Value / $maxP, 0)
        $cls = if ($entry.Key -in @("Extensions bloquees","CaractereInvalide","NomTropLong","NomReserveSharePoint")) { "seg-error" }
               elseif ($entry.Key -in @("Chemins trop longs","PrefixeInterdit","NomReserve","EspaceFin","Heritage casse")) { "seg-warn" }
               else { "seg-ok" }
        $lbl = HtmlEnc $entry.Key
        $chartProblems += "<div class='horiz-row'><div class='horiz-label' title='$lbl'>$lbl</div><div class='horiz-track'><div class='horiz-fill $cls' style='width:${pct}%;'></div></div><div class='horiz-count'>$($entry.Value)</div></div>"
    }
    if (-not $chartProblems) { $chartProblems = "<p style='color:#999;font-style:italic;'>Aucune donnee disponible.</p>" }

    # --- Top 10 dossiers N1 par volume ---
    $topFolders = @($DataInventaire |
        Where-Object { $_.TypeLigne -eq "DossierNiveau1" } |
        ForEach-Object { $_ | Add-Member -NotePropertyName _Octets -NotePropertyValue (ConvertTo-LongSafe $_.TailleOctets) -PassThru -Force } |
        Sort-Object _Octets -Descending |
        Select-Object -First 10)
    $maxF = if ($topFolders) { $topFolders[0]._Octets } else { 1 }
    if ($maxF -le 0) { $maxF = 1 }

    $chartFolders = ""
    foreach ($f in $topFolders) {
        $pct = [math]::Round(100.0 * $f._Octets / $maxF, 0)
        $lbl = if ($f.NomFileShare -and $f.NomFileShare -ne '[CONSOLIDE]') { HtmlEnc "$($f.NomFileShare)\$($f.DossierNiveau1)" } else { HtmlEnc $f.DossierNiveau1 }
        $chartFolders += "<div class='horiz-row'><div class='horiz-label'>$lbl</div><div class='horiz-track'><div class='horiz-fill seg-ok' style='width:${pct}%;'></div></div><div class='horiz-count'>$($f.TailleLisible)</div></div>"
    }
    if (-not $chartFolders) { $chartFolders = "<p style='color:#999;font-style:italic;'>Aucune donnee d'inventaire disponible.</p>" }

    return @"
<div class="section" id="section-charts">
  <h2>&#x1F4CA; Graphiques d'analyse</h2>
  <div>
    <h3>Top problemes detectes</h3>
    <div class="chart-horiz">$chartProblems</div>
  </div>
  <div style="margin-top:20px;">
    <h3>Top 10 dossiers N1 par volume</h3>
    <div class="chart-horiz">$chartFolders</div>
  </div>
</div>
"@
}

function Build-SectionActionPlan {
    param(
        [long]$NbExtensions,
        [long]$NbInvalidesError,
        [long]$NbCheminsError,
        [long]$NbVolumineuxBloquant,
        [long]$NbDoublons,
        [long]$NbPrefixesInterdits,
        [long]$NbFichiersSysteme,
        [long]$NbOrphelins      = 0,
        [double]$GoOrphelins    = 0,
        [long]$NbDesactives     = 0,
        [long]$NbInactifsHome   = 0,
        [long]$NbAbandonnes     = 0,
        [double]$GoAbandonnes   = 0,
        [long]$NbVides          = 0,
        [long]$NbAccessDenied   = 0,
        [long]$NbSidInconnus    = 0,
        [bool]$HasPermissions   = $false,
        [long]$NbRupturesHeritage = 0
    )
    $rows = [System.Collections.Generic.List[string]]::new()
    if ($NbAccessDenied -gt 0)       { $rows.Add("<tr><td class='priority-p0'>P0</td><td>&#x1F512; Dossiers refuses a l'audit</td><td><span class='badge badge-error'>ERROR</span></td><td>Decider du sort des $NbAccessDenied dossiers inaccessibles — permissions NTFS inconnues, audit incomplet.</td></tr>") }
    if ($NbCheminsError -gt 0)       { $rows.Add("<tr><td class='priority-p0'>P0</td><td>&#x1F4CF; Chemins trop longs (ERROR)</td><td><span class='badge badge-error'>ERROR</span></td><td>Raccourcir les chemins de $NbCheminsError elements depassant la limite SharePoint Online.</td></tr>") }
    if ($NbOrphelins -gt 0)          { $rows.Add("<tr><td class='priority-p0'>P0</td><td>&#x1F464; Dossiers homes orphelins</td><td><span class='badge badge-error'>ERROR</span></td><td>Decider du sort des $NbOrphelins dossiers orphelins ($([math]::Round($GoOrphelins,1)) GB) — aucun compte AD correspondant.</td></tr>") }
    if ($NbSidInconnus -gt 0)        { $rows.Add("<tr><td class='priority-p0'>P0</td><td>&#x1F464; SID de domaine inconnu dans les ACL</td><td><span class='badge badge-error'>ERROR</span></td><td>Nettoyer les $NbSidInconnus ACE d'un ancien domaine AD avant migration (SID non resolvables dans le domaine cible).</td></tr>") }
    # v3.3 — Extensions a auditer retirees du plan d'action (affichage supprime per decision client)
    if ($NbRupturesHeritage -gt 0)   { $rows.Add("<tr><td class='priority-p1'>P1</td><td>&#x1F510; Arbitrer les ruptures d'heritage NTFS</td><td><span class='badge badge-warn'>WARN</span></td><td>$NbRupturesHeritage dossiers presentent un heritage casse. Chaque cas necessite une decision avant migration : recreer les permissions explicites, creer un site/bibliotheque SharePoint dedie(e), ou aplatir l'arborescence. Voir section Permissions NTFS.</td></tr>") }
    if ($NbDesactives -gt 0)         { $rows.Add("<tr><td class='priority-p1'>P1</td><td>&#x1F464; Comptes homes desactives</td><td><span class='badge badge-warn'>WARN</span></td><td>Decider du sort des homes de $NbDesactives comptes desactives — archiver ou purger avant migration.</td></tr>") }
    if ($HasPermissions)             { $rows.Add("<tr><td class='priority-p1'>P1</td><td>&#x1F510; Mapper les groupes AD vers Entra ID</td><td><span class='badge badge-warn'>WARN</span></td><td>Etablir la correspondance groupes AD / groupes Entra ID (Microsoft 365) pour reproduire les permissions dans SPO/OneDrive.</td></tr>") }
    if ($NbInactifsHome -gt 0)       { $rows.Add("<tr><td class='priority-p1'>P1</td><td>&#x1F464; Homes comptes inactifs</td><td><span class='badge badge-warn'>WARN</span></td><td>Valider avec les managers les $NbInactifsHome homes de comptes inactifs (&gt;6 mois).</td></tr>") }
    if ($NbDoublons -gt 0)           { $rows.Add("<tr><td class='priority-p2'>P2</td><td>&#x1F4D1; Doublons</td><td><span class='badge badge-warn'>WARN</span></td><td>Deduplication de $NbDoublons fichiers dupliques — reduire le volume et eviter les conflits.</td></tr>") }
    if ($NbPrefixesInterdits -gt 0)  { $rows.Add("<tr><td class='priority-p2'>P2</td><td>&#x1F524; Prefixes # / %</td><td><span class='badge badge-warn'>WARN</span></td><td>Renommer les $NbPrefixesInterdits elements avec des prefixes non supportes par SharePoint Online.</td></tr>") }
    if ($NbAbandonnes -gt 0)         { $rows.Add("<tr><td class='priority-p2'>P2</td><td>&#x1F464; Homes abandonnes</td><td><span class='badge badge-info'>INFO</span></td><td>Purger les $NbAbandonnes dossiers homes abandonnes ($([math]::Round($GoAbandonnes,1)) GB recup.).</td></tr>") }
    if ($NbFichiersSysteme -gt 0)    { $rows.Add("<tr><td class='priority-p3'>P3</td><td>&#x1F5D1; Fichiers systeme/temporaires</td><td><span class='badge badge-info'>INFO</span></td><td>Les $NbFichiersSysteme fichiers systeme (desktop.ini, ~`$, .lock) sont skipped automatiquement par ShareGate — aucune action requise.</td></tr>") }
    if ($NbVides -gt 0)              { $rows.Add("<tr><td class='priority-p3'>P3</td><td>&#x1F464; Dossiers homes vides/techniques</td><td><span class='badge badge-info'>INFO</span></td><td>Exclure les $NbVides dossiers vides ou techniques de la migration OneDrive.</td></tr>") }

    if ($rows.Count -eq 0) {
        $rows.Add("<tr><td colspan='4' style='text-align:center;color:#27ae60;padding:20px;'>&#x2705; Aucune action corrective requise — migration prete !</td></tr>")
    }

    return @"
<div class="section" id="section-plan">
  <h2>&#x1F3AF; Plan d'action priorise</h2>
  <p style="margin-bottom:12px;font-size:13px;color:#555;">
    <strong>P0</strong> : Bloquant (a traiter avant migration) &nbsp;|&nbsp;
    <strong>P1</strong> : Critique (a traiter en priorite) &nbsp;|&nbsp;
    <strong>P2</strong> : Important &nbsp;|&nbsp;
    <strong>P3</strong> : Optimisation
  </p>
  <table class="plan-table">
    <tr><th>Priorite</th><th>Action</th><th>Severite</th><th>Detail</th></tr>
    $($rows -join "`n    ")
  </table>
</div>
"@
}

function Build-SectionInventaire {
    param([array]$DataInventaire)
    if ($DataInventaire.Count -eq 0) { return "" }

    $ligneTotal  = $DataInventaire | Where-Object {
        $_.DossierNiveau1 -eq "[TOTAL]" -and (
            -not $_.PSObject.Properties['NomFileShare'] -or $_.NomFileShare -eq "[CONSOLIDE]"
        )
    } | Select-Object -First 1
    if (-not $ligneTotal) {
        $ligneTotal = $DataInventaire | Where-Object { $_.DossierNiveau1 -eq "[TOTAL]" } | Select-Object -First 1
    }
    $totalOctets = if ($ligneTotal) { ConvertTo-LongSafe $ligneTotal.TailleOctets } else { 0 }

    $n1Rows = @($DataInventaire | Where-Object { $_.TypeLigne -eq "DossierNiveau1" })

    if ($n1Rows.Count -gt 0) {
        # Standard case: list N1 rows then the [TOTAL] row
        $inventaireRows = $n1Rows + @(if ($ligneTotal) { $ligneTotal } else { @() })
    } elseif ($ligneTotal) {
        # Share with files only at root level (no N1 sub-folders) — show a [Racine] synthetic row
        $racineRow = [PSCustomObject]@{
            TypeLigne      = "DossierNiveau1"
            DossierNiveau1 = "[Racine]"
            NombreFichiers = $ligneTotal.NombreFichiers
            NombreDossiers = $ligneTotal.NombreDossiers
            TailleOctets   = $ligneTotal.TailleOctets
            TailleLisible  = $ligneTotal.TailleLisible
            NomFileShare   = ""
        }
        $inventaireRows = @($racineRow, $ligneTotal)
    } else {
        $inventaireRows = @()
    }

    $rows = [System.Collections.Generic.List[string]]::new()
    foreach ($ligne in $inventaireRows) {
        if ($ligne.DossierNiveau1 -eq "[TOTAL]") {
            $rows.Add("<tr class='tbl-total'><td><strong>[TOTAL]</strong></td><td><strong>$($ligne.NombreFichiers)</strong></td><td><strong>$($ligne.NombreDossiers)</strong></td><td><strong>$($ligne.TailleLisible)</strong></td><td><strong>100%</strong></td></tr>")
        } else {
            $octetsLigne = ConvertTo-LongSafe $ligne.TailleOctets
            $pct = if ($totalOctets -gt 0) { [math]::Round(100.0 * $octetsLigne / $totalOctets, 1) } else { 0 }
            $volBadge = if ($octetsLigne -gt ($script:VOL_ELEVE_TB * 1099511627776)) { "<span class='badge-volume'>&#x26A0; Volume eleve</span>" } else { "" }
            $fsPrefix = if ($ligne.PSObject.Properties['NomFileShare'] -and $ligne.NomFileShare -and $ligne.NomFileShare -ne '[CONSOLIDE]') { "$(HtmlEnc $ligne.NomFileShare)\" } else { "" }
            $n1 = HtmlEnc $ligne.DossierNiveau1
            $rows.Add("<tr><td>$fsPrefix$n1$volBadge</td><td>$($ligne.NombreFichiers)</td><td>$($ligne.NombreDossiers)</td><td>$($ligne.TailleLisible)</td><td>${pct}%</td></tr>")
        }
    }

    return @"
<div class="section" id="section-inventaire">
  <h2>&#x1F4E6; Inventaire du FileShare</h2>
  <table>
    <tr><th>Dossier Niveau 1</th><th>Fichiers</th><th>Dossiers</th><th>Volume</th><th>% du total</th></tr>
    $($rows -join "`n    ")
  </table>
</div>
"@
}

function Build-DetailTable {
    param(
        [string]$SectionId,
        [string]$Title,
        [string]$Counter,
        [string]$HeaderRow,
        [System.Collections.Generic.List[string]]$DataRows,
        [string]$CsvFileName = ""
    )
    if ($DataRows.Count -eq 0) { return "" }

    $total     = $DataRows.Count
    $displayed = $DataRows | Select-Object -First $MAX_ROWS
    $remainder = $total - ($displayed | Measure-Object).Count

    $footer = ""
    if ($remainder -gt 0) {
        $fn = if ($CsvFileName) { " Consultez le fichier CSV complet : <code>$CsvFileName</code>" } else { "" }
        $footer = "<tr class='table-footer'><td colspan='99'>... et $remainder autres elements.$fn</td></tr>"
    }

    return @"
<div class="section" id="$SectionId">
  <h2>$Title</h2>
  <div class="section-counter">$Counter</div>
  <table>
    <tr>$HeaderRow</tr>
    $($displayed -join "`n    ")
    $footer
  </table>
</div>
"@
}

function Build-SectionVolumineux {
    param([array]$DataVolumineux, [string]$CsvName = "")
    if ($DataVolumineux.Count -eq 0) { return "" }
    $nbError = ($DataVolumineux | Where-Object { $_.Severite -eq "ERROR" } | Measure-Object).Count
    $nbWarn  = ($DataVolumineux | Where-Object { $_.Severite -eq "WARN"  } | Measure-Object).Count
    $counter = "<span class='badge badge-error'>$nbError BLOQUANTS</span>&nbsp;<span class='badge badge-warn'>$nbWarn a surveiller</span>"
    $header  = "<th>Fichier</th><th>Taille</th><th>Categorie</th><th>Severite</th><th>Derniere modification</th><th>Emplacement</th>"
    $rows = [System.Collections.Generic.List[string]]::new()
    foreach ($r in ($DataVolumineux | Sort-Object { [double](($_.TailleGB -replace ',','.') -replace '[^0-9.]','0') } -Descending)) {
        $taille = if ([double](($r.TailleGB -replace ',','.') -replace '[^0-9.]','0') -ge 1) { "$($r.TailleGB) GB" } else { "$($r.TailleMB) MB" }
        $badge  = if ($r.Severite -eq "ERROR") { "<span class='badge badge-error'>ERROR</span>" } else { "<span class='badge badge-warn'>WARN</span>" }
        $empl   = Get-EmplacementCell $r.CheminComplet
        $rows.Add("<tr><td>$(HtmlEnc $r.NomFichier)</td><td>$taille</td><td>$(HtmlEnc $r.Categorie)</td><td>$badge</td><td>$(HtmlEnc $r.DerniereModification)</td>$empl</tr>")
    }
    return Build-DetailTable -SectionId "section-volumineux" -Title "&#x26A0;&#xFE0F; Fichiers Volumineux" -Counter $counter -HeaderRow $header -DataRows $rows -CsvFileName $CsvName
}

function Build-SectionInvalides {
    param([array]$DataInvalides, [string]$CsvName = "")
    if ($DataInvalides.Count -eq 0) { return "" }
    $nbError = ($DataInvalides | Where-Object { $_.Severite -eq "ERROR" } | Measure-Object).Count
    $nbWarn  = ($DataInvalides | Where-Object { $_.Severite -eq "WARN"  } | Measure-Object).Count
    $nbInfo  = ($DataInvalides | Where-Object { $_.Severite -eq "INFO"  } | Measure-Object).Count
    $counter = "<span class='badge badge-error'>$nbError erreurs</span>&nbsp;<span class='badge badge-warn'>$nbWarn warnings</span>&nbsp;<span class='badge badge-info'>$nbInfo info</span>"
    $header  = "<th>Nom</th><th>Type</th><th>Probleme</th><th>Severite</th><th>Detail</th><th>Emplacement</th>"
    $rows = [System.Collections.Generic.List[string]]::new()
    foreach ($r in ($DataInvalides | Sort-Object Severite)) {
        $badge = switch ($r.Severite) {
            "ERROR" { "<span class='badge badge-error'>ERROR</span>" }
            "WARN"  { "<span class='badge badge-warn'>WARN</span>" }
            "INFO"  { "<span class='badge badge-info'>INFO</span>" }
            default { "<span class='badge badge-warn'>WARN</span>" }
        }
        $empl  = Get-EmplacementCell $r.CheminComplet
        $rows.Add("<tr><td>$(HtmlEnc $r.Nom)</td><td>$(HtmlEnc $r.TypeElement)</td><td>$(HtmlEnc $r.ProblemeType)</td><td>$badge</td><td>$(HtmlEnc $r.Detail)</td>$empl</tr>")
    }
    return Build-DetailTable -SectionId "section-invalides" -Title "&#x274C; Caracteres Invalides pour M365" -Counter $counter -HeaderRow $header -DataRows $rows -CsvFileName $CsvName
}

function Build-SectionChemins {
    param([array]$DataChemins, [string]$CsvName = "")
    if ($DataChemins.Count -eq 0) { return "" }
    $counter = "<span class='badge badge-warn'>$($DataChemins.Count) chemins trop longs</span>"
    $header  = "<th>Nom</th><th>Type</th><th>Longueur</th><th>Depassement</th><th>Emplacement</th>"
    $rows = [System.Collections.Generic.List[string]]::new()
    foreach ($r in ($DataChemins | Sort-Object { [int](($_.Depassement -replace '\D','') -replace '^$','0') } -Descending)) {
        $empl = Get-EmplacementCell $r.CheminComplet
        $rows.Add("<tr><td>$(HtmlEnc $r.Nom)</td><td>$(HtmlEnc $r.TypeElement)</td><td>$($r.LongueurChemin)</td><td><span class='badge badge-warn'>+$($r.Depassement) car.</span></td>$empl</tr>")
    }
    return Build-DetailTable -SectionId "section-chemins" -Title "&#x1F4CF; Chemins Trop Longs" -Counter $counter -HeaderRow $header -DataRows $rows -CsvFileName $CsvName
}

function Build-SectionDoublons {
    param([array]$DataDoublons, [string]$CsvName = "")
    if ($DataDoublons.Count -eq 0) { return "" }

    # Grouper par GroupeDoublonId
    $groupes = @($DataDoublons | Group-Object GroupeDoublonId)

    # Stats globales
    $nbGroupes  = $groupes.Count
    $nbFichiers = $DataDoublons.Count

    # Volume récupérable = somme sur chaque groupe de (NombreOccurrences - 1) x TailleOctets
    $volRecupOctets = [long]0
    $plusGrosTailleMB = 0.0
    foreach ($g in $groupes) {
        $fr  = $g.Group[0]
        $occ = [int](($fr.NombreOccurrences -replace '[^\d]','') -replace '^$','0')
        $tailleOctets = ConvertTo-LongSafe $fr.TailleOctets
        $tailleMB = [double](($fr.TailleMB -replace ',','.') -replace '[^0-9.]','0')
        if ($occ -gt 1) { $volRecupOctets += ($occ - 1) * $tailleOctets }
        if ($tailleMB -gt $plusGrosTailleMB) { $plusGrosTailleMB = $tailleMB }
    }
    $volRecupGB = [math]::Round($volRecupOctets / 1073741824.0, 2)
    $volRecupMB = [math]::Round($volRecupOctets / 1048576.0, 1)
    $volRecupStr  = if ($volRecupGB -ge 1) { "$volRecupGB GB" } else { "$volRecupMB MB" }
    $plusGrosStr  = "$([math]::Round($plusGrosTailleMB, 1)) MB"

    # Construire les données par groupe pour le tri Top 20
    $groupesData = [System.Collections.Generic.List[object]]::new()
    foreach ($g in $groupes) {
        $fr   = $g.Group[0]
        $occ  = [int](($fr.NombreOccurrences -replace '\D','') -replace '^$','0')
        $tailleMB = [double](($fr.TailleMB -replace ',','.') -replace '[^0-9.]','0')
        $volRecup = if ($occ -gt 1) { [math]::Round(($occ - 1) * $tailleMB, 2) } else { 0.0 }
        $groupesData.Add([PSCustomObject]@{
            NomFichier  = $fr.NomFichier
            Occurrences = $occ
            TailleMB    = $tailleMB
            VolRecupMB  = $volRecup
            Chemins     = @($g.Group | Select-Object -ExpandProperty CheminComplet)
        })
    }

    $top20     = @($groupesData | Sort-Object VolRecupMB -Descending | Select-Object -First 20)
    $remainder = $groupesData.Count - $top20.Count

    $rows = [System.Collections.Generic.List[string]]::new()
    $i = 0
    foreach ($g in $top20) {
        $i++
        $tU = if ($g.TailleMB -ge 1024) { "$([math]::Round($g.TailleMB/1024,2)) GB" } else { "$([math]::Round($g.TailleMB,1)) MB" }
        $vR = if ($g.VolRecupMB -ge 1024) { "$([math]::Round($g.VolRecupMB/1024,2)) GB" } else { "$([math]::Round($g.VolRecupMB,1)) MB" }
        $cheminsList = ($g.Chemins | ForEach-Object { "<li style='word-break:break-all;'>$(HtmlEnc $_)</li>" }) -join ""
        $emplCell = "<td class='path-cell'><details><summary style='cursor:pointer;color:#2980b9;'>$($g.Chemins.Count) emplacement(s)</summary><ul style='margin:4px 0 0 12px;font-size:11px;list-style:disc;'>$cheminsList</ul></details></td>"
        $rows.Add("<tr><td>$i</td><td>$(HtmlEnc $g.NomFichier)</td><td>$($g.Occurrences)</td><td>$tU</td><td>$vR</td>$emplCell</tr>")
    }

    $footer = ""
    if ($remainder -gt 0) {
        $fn = if ($CsvName) { " Consultez le CSV : <code>$CsvName</code>" } else { "" }
        $footer = "<tr class='table-footer'><td colspan='6'>... et $remainder autres groupes.$fn</td></tr>"
    }

    return @"
<div class="section" id="section-doublons">
  <h2>&#x1F4D1; Fichiers dupliques</h2>
  <div class="stat-row">
    <div class="stat-box"><div class="number">$nbGroupes</div><div class="label">Groupes de doublons</div></div>
    <div class="stat-box warn"><div class="number">$nbFichiers</div><div class="label">Fichiers dupliques</div></div>
    <div class="stat-box ok"><div class="number">$volRecupStr</div><div class="label">Volume recuperable estime</div></div>
    <div class="stat-box"><div class="number">$plusGrosStr</div><div class="label">Plus gros doublon</div></div>
  </div>
  <table>
    <tr><th>#</th><th>Nom du fichier</th><th>Occurrences</th><th>Taille unitaire</th><th>Volume recuperable</th><th>Emplacements</th></tr>
    $($rows -join "`n    ")
    $footer
  </table>
</div>
"@
}

function Build-SectionExtensions {
    param([array]$DataExtensions, [string]$CsvName = "")
    if ($DataExtensions.Count -eq 0) { return "" }
    $counter = "<span class='badge badge-warn'>$($DataExtensions.Count) fichiers a auditer</span>"
    $header  = "<th>Fichier</th><th>Extension</th><th>Taille (MB)</th><th>Severite</th><th>Emplacement</th>"
    $rows = [System.Collections.Generic.List[string]]::new()
    foreach ($r in ($DataExtensions | Sort-Object Extension)) {
        $badge = if ($r.Severite -eq "ERROR") { "<span class='badge badge-error'>ERROR</span>" } else { "<span class='badge badge-warn'>WARN</span>" }
        $empl = Get-EmplacementCell $r.CheminComplet
        $rows.Add("<tr><td>$(HtmlEnc $r.NomFichier)</td><td><strong>$(HtmlEnc $r.Extension)</strong></td><td>$($r.TailleMB)</td><td>$badge</td>$empl</tr>")
    }
    return Build-DetailTable -SectionId "section-extensions" -Title "&#x26A0; Extensions a Auditer (blocklist historique SP Server)" -Counter $counter -HeaderRow $header -DataRows $rows -CsvFileName $CsvName
}

function Build-SectionPermissions {
    param(
        [array]$DataPermissions,
        [string]$CsvName = "",
        [array]$DataSidResolution = @(),
        [array]$DataAccessDenied  = @(),
        [string]$NomFS      = "",
        [string]$OutputPath = "",
        [string]$Timestamp  = "",
        [string]$CheminUNCRoot = "",
        [bool]$PermissionsCsvAvailable = $false
    )
    $hasPermissionsSection = $PermissionsCsvAvailable -or ($DataPermissions.Count -gt 0)
    if (-not $hasPermissionsSection) { return "" }
    if ($DataPermissions.Count -eq 0) {
        $scopeInfo = if (-not [string]::IsNullOrWhiteSpace($CheminUNCRoot)) { "<code>$(HtmlEnc $CheminUNCRoot)</code>" } else { "ce FileShare" }
        return @"
<div class="section" id="section-permissions">
  <h2>&#x1F510; Permissions NTFS</h2>
  <div class="info-block" style="background:#fef9e7;border-left:4px solid #e67e22;padding:12px 16px;margin:8px 0;border-radius:4px;">
    <strong>&#x26A0; Aucune permission NTFS détectée pour ce FileShare.</strong>
    Vérifier que <code>Get-NTFSPermissions.ps1</code> a été exécuté sur $scopeInfo.
  </div>
</div>
"@
    }
    $nb = $DataPermissions.Count

    # Stats globales
    $nbHeritage = ($DataPermissions | Where-Object { $_.EstHerite -eq "True" -or $_.EstHerite -eq "1" } | Measure-Object).Count
    $pctHeritage = if ($nb -gt 0) { [math]::Round(100.0 * $nbHeritage / $nb, 1) } else { 0 }
    $nbExplicite = $nb - $nbHeritage

    # Identités uniques
    $idCol = if ($DataPermissions[0].PSObject.Properties.Name -contains "Identite") { "Identite" } else { "IdentityReference" }
    $identites = @($DataPermissions | Select-Object -ExpandProperty $idCol -Unique | Where-Object { $_ })
    $nbIdentites = $identites.Count

    # Ruptures d'héritage (ACE explicites)
    $cheminCol = if ($DataPermissions[0].PSObject.Properties.Name -contains "CheminDossier") { "CheminDossier" } else { "Path" }
    $dossiersRupture = @($DataPermissions | Where-Object { $_.EstHerite -ne "True" -and $_.EstHerite -ne "1" } | Select-Object -ExpandProperty $cheminCol -Unique | Where-Object { $_ })
    $nbRuptures = $dossiersRupture.Count

    # ----------------------------------------------------------------
    # HOTSPOT ANALYSIS — ruptures d'heritage
    # ----------------------------------------------------------------
    $hotspotHtml    = ""
    $csvHotspotsName = ""
    if ($nbRuptures -gt 0) {
        $ruptureAces = $DataPermissions | Where-Object { $_.EstHerite -ne "True" -and $_.EstHerite -ne "1" }
        $hotspotData = [System.Collections.Generic.List[object]]::new()
        foreach ($fg in ($ruptureAces | Group-Object $cheminCol)) {
            $chemin = $fg.Name
            if ([string]::IsNullOrWhiteSpace($chemin)) { continue }
            $nbAce    = $fg.Count
            $nbIdUniq = ($fg.Group | Select-Object -ExpandProperty $idCol -Unique | Where-Object { $_ }).Count
            $proprio  = ($fg.Group | Where-Object { $_.Proprietaire } | Select-Object -ExpandProperty "Proprietaire" -First 1)
            if (-not $proprio) { $proprio = "" }
            $profondeur = if (-not [string]::IsNullOrWhiteSpace($CheminUNCRoot)) {
                Get-DepthRelativeToShareRoot -Path $chemin -ShareRoot $CheminUNCRoot
            } else {
                ($chemin.TrimStart('\', '/') -split '[/\\]' | Where-Object { $_ }).Count
            }
            $hotspotData.Add([PSCustomObject]@{
                CheminDossier      = $chemin
                Profondeur         = $profondeur
                NbACEExplicites    = $nbAce
                NbIdentitesUniques = $nbIdUniq
                Proprietaire       = $proprio
            })
        }
        $hotspotsSorted = @($hotspotData | Sort-Object @{Expression="NbACEExplicites";Descending=$true}, @{Expression="Profondeur";Descending=$false})

        # N1 breakdown — v3.3 patch: use Get-N1RelativeToShareRoot helper
        $n1BrokenCounts = @{}
        foreach ($h in $hotspotsSorted) {
            $chemin = $h.CheminDossier
            $n1 = if (-not [string]::IsNullOrWhiteSpace($CheminUNCRoot)) {
                $r = Get-N1RelativeToShareRoot -Path $chemin -ShareRoot $CheminUNCRoot
                if ([string]::IsNullOrWhiteSpace($r)) { "[Racine]" } else { $r }
            } else {
                # Fallback: strip server+share (2 UNC components) then get N1
                $parts2 = $chemin.TrimStart('\', '/') -split '[/\\]' | Where-Object { $_ }
                if ($parts2.Count -gt 2) { $parts2[2] } elseif ($parts2.Count -gt 1) { $parts2[1] } elseif ($parts2.Count -gt 0) { $parts2[0] } else { "" }
            }
            if (-not [string]::IsNullOrWhiteSpace($n1)) {
                if (-not $n1BrokenCounts.ContainsKey($n1)) { $n1BrokenCounts[$n1] = 0 }
                $n1BrokenCounts[$n1]++
            }
        }
        $nbN1Impactes = $n1BrokenCounts.Count

        # % arborescence
        $nbDossiersTotal = ($DataPermissions | Select-Object -ExpandProperty $cheminCol -Unique | Where-Object { $_ }).Count
        $pctRuptures = if ($nbDossiersTotal -gt 0) { [math]::Round(100.0 * $nbRuptures / $nbDossiersTotal, 1) } else { 0 }

        # Max depth of a rupture
        $profondeurMax = if ($hotspotsSorted.Count -gt 0) { ($hotspotsSorted | Measure-Object Profondeur -Maximum).Maximum } else { 0 }

        # Generate InheritanceHotspots CSV
        $csvHotspotsName = if (-not [string]::IsNullOrWhiteSpace($NomFS)) { "InheritanceHotspots_${NomFS}.csv" } else { "InheritanceHotspots.csv" }
        if (-not [string]::IsNullOrWhiteSpace($OutputPath) -and -not [string]::IsNullOrWhiteSpace($NomFS) -and (Test-Path $OutputPath)) {
            $ts = if ($Timestamp) { $Timestamp } else { Get-Date -Format "yyyyMMdd_HHmmss" }
            $csvHotPath = Join-Path $OutputPath "InheritanceHotspots_${NomFS}_${ts}.csv"
            try {
                $hotspotsSorted | Select-Object CheminDossier, Profondeur, NbACEExplicites, NbIdentitesUniques, Proprietaire |
                    Export-Csv -Path $csvHotPath -Delimiter ";" -Encoding UTF8 -NoTypeInformation -Force
                $csvHotspotsName = Split-Path -Leaf $csvHotPath
            } catch {
                Write-Log "Avertissement : impossible de generer InheritanceHotspots CSV : $_" "WARN"
            }
        }

        # Top 10 N1 chart
        $top10N1  = $n1BrokenCounts.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 10
        $maxN1Val = if ($top10N1) { ($top10N1 | Measure-Object Value -Maximum).Maximum } else { 1 }
        if ($maxN1Val -le 0) { $maxN1Val = 1 }
        $chartN1 = ""
        foreach ($entry in $top10N1) {
            $pctN1 = [math]::Round(100.0 * $entry.Value / $maxN1Val, 0)
            $lblN1 = HtmlEnc $entry.Key
            $chartN1 += "<div class='horiz-row'><div class='horiz-label'>$lblN1</div><div class='horiz-track'><div class='horiz-fill seg-warn' style='width:${pctN1}%;'></div></div><div class='horiz-count'>$($entry.Value)</div></div>"
        }

        # Table top 50 hotspots
        $MAX_HOT = 50
        $hotTblRows = [System.Collections.Generic.List[string]]::new()
        foreach ($h in ($hotspotsSorted | Select-Object -First $MAX_HOT)) {
            $hotTblRows.Add("<tr><td class='path-cell'>$(HtmlEnc $h.CheminDossier)</td><td>$($h.Profondeur)</td><td>$($h.NbACEExplicites)</td><td>$($h.NbIdentitesUniques)</td><td>$(HtmlEnc $h.Proprietaire)</td></tr>")
        }
        $hotFooter = if ($hotspotsSorted.Count -gt $MAX_HOT) { "<tr class='table-footer'><td colspan='5'>... et $($hotspotsSorted.Count - $MAX_HOT) autres elements. Consultez le fichier CSV complet : <code>$csvHotspotsName</code></td></tr>" } else { "" }

        $hotspotHtml = @"
  <h3>Hotspots de ruptures d'heritage</h3>
  <div class="info-block" style="background:#eaf4fb;border-left:4px solid #2980b9;padding:12px 16px;margin:8px 0;border-radius:4px;">
    <strong>&#x2139; Pourquoi c'est critique ?</strong> ShareGate migre les permissions explicites mais SharePoint Online gere mal les arborescences profondes avec ruptures d'heritage. Chaque rupture necessite un arbitrage avant migration : recreer la permission, creer un site/bibliotheque dedie, ou aplatir l'arborescence.
  </div>
  <div class="stat-row">
    <div class="stat-box warn"><div class="number">$nbRuptures</div><div class="label">Dossiers heritage casse</div></div>
    <div class="stat-box warn"><div class="number">$nbN1Impactes</div><div class="label">Dossiers N1 impactes</div></div>
    <div class="stat-box warn"><div class="number">$pctRuptures%</div><div class="label">% arborescence concernee</div></div>
    <div class="stat-box"><div class="number">$profondeurMax</div><div class="label">Profondeur max rupture</div></div>
  </div>
  <h3>Top 10 dossiers N1 par nb de ruptures</h3>
  <div class="chart-horiz">$chartN1</div>
  <h3>Top $MAX_HOT dossiers avec heritage casse</h3>
  <table>
    <tr><th>Chemin du dossier</th><th>Profondeur</th><th>Nb ACE explicites</th><th>Nb identites uniques</th><th>Proprietaire</th></tr>
    $($hotTblRows -join "`n    ")
    $hotFooter
  </table>
"@
    }

    # ----------------------------------------------------------------
    # IDENTITY INVENTORY
    # ----------------------------------------------------------------
    $identitiesHtml = ""
    if ($DataPermissions.Count -gt 0) {
        $identityInventory = [System.Collections.Generic.List[object]]::new()
        foreach ($ig in ($DataPermissions | Group-Object $idCol)) {
            $idRef = $ig.Name
            if ([string]::IsNullOrWhiteSpace($idRef)) { continue }
            $idType = "User"; $idDomaine = ""
            if ($idRef -match '^S-\d+-') {
                $idType = "SID brut"; $idDomaine = ""
            } elseif ($idRef -match '^(?i)(BUILTIN|NT AUTHORITY|NT SERVICE)\\') {
                $idType = "BUILTIN"; $idDomaine = ($idRef -replace '\\.*$','')
            } elseif ($idRef -match '^([^\\]+)\\(.+)$') {
                $idDomaine = $Matches[1]; $idName2 = $Matches[2]
                # v3.3 — Heuristique elargie pour groupes AD PrimaGAZ (ordre de priorite)
                # 1. Groupes bien connus (noms generiques de domaine)
                $wellKnownGroups = @(
                    'Domain Admins','Domain Users','Domain Computers','Domain Guests',
                    'Enterprise Admins','Schema Admins','Group Policy Creator Owners',
                    'Administrators','Users','Everyone','Authenticated Users'
                )
                if ($idName2 -in $wellKnownGroups) {
                    $idType = "Group"
                # 2. Domaine SHVE avec nomenclature USR_BTQ (groupes PrimaGAZ)
                } elseif ($idDomaine -ieq 'SHVE' -and $idName2 -match '(?i)USR_BTQ') {
                    $idType = "Group"
                # 3. Pattern _G_ ou _U_ entre underscores (USR_BTQ_G_*, USR_BTQ_U_*)
                } elseif ($idName2 -match '(?i)_[GU]_') {
                    $idType = "Group"
                # 4. Regex historique
                } elseif ($idName2 -match '(?i)(_Group_|_Grp_|_GG_|_DL_|_GL_|FR_PG_|^GG_|^DL_|^GL_|^GRP_|^GRPE)') {
                    $idType = "Group"
                }
            } elseif ($idRef -match '^(?i)(Everyone|CREATOR OWNER|Authenticated Users)$') {
                $idType = "BUILTIN"
            }
            $identityInventory.Add([PSCustomObject]@{
                Identite      = $idRef
                Type          = $idType
                Domaine       = $idDomaine
                NbOccurrences = $ig.Count
                NbDossiers    = ($ig.Group | Select-Object -ExpandProperty $cheminCol -Unique | Where-Object { $_ }).Count
            })
        }
        $idSorted = @($identityInventory | Sort-Object @{
            Expression = {
                switch ($_.Type) {
                    "Group"    { 0 }
                    "User"     { 1 }
                    "BUILTIN"  { 2 }
                    "SID brut" { 3 }
                    default    { 4 }
                }
            }
        }, @{Expression="NbOccurrences";Descending=$true})
        $nbTotalIds = $idSorted.Count
        $nbGroups   = ($idSorted | Where-Object { $_.Type -eq "Group"  } | Measure-Object).Count
        $nbUsers    = ($idSorted | Where-Object { $_.Type -eq "User"   } | Measure-Object).Count

        # Generate IdentitiesInventory CSV
        $csvIdName = if (-not [string]::IsNullOrWhiteSpace($NomFS)) { "IdentitiesInventory_${NomFS}.csv" } else { "IdentitiesInventory.csv" }
        if (-not [string]::IsNullOrWhiteSpace($OutputPath) -and -not [string]::IsNullOrWhiteSpace($NomFS) -and (Test-Path $OutputPath)) {
            $ts2 = if ($Timestamp) { $Timestamp } else { Get-Date -Format "yyyyMMdd_HHmmss" }
            $csvIdPath = Join-Path $OutputPath "IdentitiesInventory_${NomFS}_${ts2}.csv"
            try {
                $idSorted | Select-Object Identite, Type, Domaine, NbOccurrences, NbDossiers |
                    Export-Csv -Path $csvIdPath -Delimiter ";" -Encoding UTF8 -NoTypeInformation -Force
                $csvIdName = Split-Path -Leaf $csvIdPath
            } catch {
                Write-Log "Avertissement : impossible de generer IdentitiesInventory CSV : $_" "WARN"
            }
        }

        $idTblRows = [System.Collections.Generic.List[string]]::new()
        foreach ($id in $idSorted) {
            $typeBadge = switch ($id.Type) {
                "Group"    { "<span class='badge badge-warn'>Group</span>" }
                "User"     { "<span class='badge badge-ok'>User</span>" }
                "BUILTIN"  { "<span class='badge badge-info'>BUILTIN</span>" }
                "SID brut" { "<span class='badge badge-error'>SID brut</span>" }
                default    { "<span class='badge'>$(HtmlEnc $id.Type)</span>" }
            }
            $idTblRows.Add("<tr><td>$(HtmlEnc $id.Identite)</td><td>$typeBadge</td><td>$(HtmlEnc $id.Domaine)</td><td>$($id.NbOccurrences)</td><td>$($id.NbDossiers)</td></tr>")
        }
        $csvIdRef = if ($csvIdName) { "<p style='font-size:12px;color:#666;'>Fichier CSV complet : <code>$csvIdName</code></p>" } else { "" }

        $identitiesHtml = @"
  <h3>Inventaire des identites AD</h3>
  <div class="info-block" style="background:#eaf4fb;border-left:4px solid #2980b9;padding:12px 16px;margin:8px 0;border-radius:4px;">
    <strong>&#x2139; A verifier dans Entra ID</strong> &mdash; Liste exhaustive des groupes AD et comptes presents dans les ACL de ce FileShare. Chaque entree doit etre verifiee dans Entra ID (Microsoft 365) pour s'assurer qu'une correspondance existe avant la migration vers SPO/OneDrive.
  </div>
  <div class="stat-row">
    <div class="stat-box"><div class="number">$nbTotalIds</div><div class="label">Identites uniques</div></div>
    <div class="stat-box warn"><div class="number">$nbGroups</div><div class="label">Groupes AD</div></div>
    <div class="stat-box ok"><div class="number">$nbUsers</div><div class="label">Comptes utilisateurs</div></div>
  </div>
  $csvIdRef
  <table>
    <tr><th>Identite</th><th>Type</th><th>Domaine</th><th>Nb occurrences (ACE)</th><th>Nb dossiers</th></tr>
    $($idTblRows -join "`n    ")
  </table>
"@
    }

    # Stat SID résolution
    $sidSection = ""
    if ($DataSidResolution.Count -gt 0) {
        $nbSidTotal    = $DataSidResolution.Count
        $nbResolu      = ($DataSidResolution | Where-Object { $_.Statut -eq "RESOLU" }          | Measure-Object).Count
        $nbDomInc      = ($DataSidResolution | Where-Object { $_.Statut -eq "DOMAINE_INCONNU" } | Measure-Object).Count
        $nbMachLoc     = ($DataSidResolution | Where-Object { $_.Statut -eq "MACHINE_LOCALE" }  | Measure-Object).Count
        $nbOrphelin    = ($DataSidResolution | Where-Object { $_.Statut -eq "ORPHELIN_VRAI" }   | Measure-Object).Count
        $alerteDomaine = if ($nbDomInc -gt 0) { "<div style='background:#fef9e7;border:1px solid #e67e22;padding:10px 14px;border-radius:6px;margin:12px 0;'><strong>&#x26A0; $nbDomInc SID de domaine(s) inconnu(s) detecte(s)</strong> — Probable ancien domaine dans les ACL (ex: NTX dans environnement SHVE). Nettoyer avant migration.</div>" } else { "" }
        $sidRows = [System.Collections.Generic.List[string]]::new()
        foreach ($r in ($DataSidResolution | Sort-Object Statut, NbOccurrencesACL -Descending | Select-Object -First 30)) {
            $badgeSid = switch ($r.Statut) {
                "RESOLU"          { "<span class='badge badge-ok'>RESOLU</span>" }
                "DOMAINE_INCONNU" { "<span class='badge badge-warn'>DOMAINE_INCONNU</span>" }
                "MACHINE_LOCALE"  { "<span class='badge badge-info'>MACHINE_LOCALE</span>" }
                "ORPHELIN_VRAI"   { "<span class='badge badge-error'>ORPHELIN_VRAI</span>" }
                default           { "<span class='badge badge-warn'>$($r.Statut)</span>" }
            }
            $nomDisplay = if ($r.Nom) { HtmlEnc $r.Nom } else { HtmlEnc $r.SID }
            $sidRows.Add("<tr><td style='font-size:11px;word-break:break-all;'>$(HtmlEnc $r.SID)</td><td>$(HtmlEnc $r.PrefixeDomaine)</td><td>$badgeSid</td><td>$nomDisplay</td><td>$($r.NbOccurrencesACL)</td></tr>")
        }
        $footerSid = if ($DataSidResolution.Count -gt 30) { "<tr class='table-footer'><td colspan='5'>... et $($DataSidResolution.Count - 30) autres. Consultez : <code>$(if($csvSidResolution){$csvSidResolution.Name})</code></td></tr>" } else { "" }
        $sidSection = @"
  $alerteDomaine
  <h3>Identites &amp; SID ($nbSidTotal uniques)</h3>
  <div class="stat-row">
    <div class="stat-box ok"><div class="number">$nbResolu</div><div class="label">Resolus</div></div>
    <div class="stat-box warn"><div class="number">$nbDomInc</div><div class="label">Domaine inconnu</div></div>
    <div class="stat-box"><div class="number">$nbMachLoc</div><div class="label">Machine locale</div></div>
    <div class="stat-box danger"><div class="number">$nbOrphelin</div><div class="label">Orphelins vrais</div></div>
  </div>
  <table>
    <tr><th>SID</th><th>Prefixe domaine</th><th>Statut</th><th>Nom</th><th>Nb ACL</th></tr>
    $($sidRows -join "`n    ")
    $footerSid
  </table>
"@
    }

    # AccessDenied
    $accessDeniedSection = ""
    if ($DataAccessDenied.Count -gt 0) {
        $adRows = [System.Collections.Generic.List[string]]::new()
        foreach ($r in ($DataAccessDenied | Select-Object -First 20)) {
            $cheminAD = if ($r.Chemin) { HtmlEnc $r.Chemin } elseif ($r.CheminDossier) { HtmlEnc $r.CheminDossier } else { "—" }
            $adRows.Add("<tr><td class='path-cell'>$cheminAD</td><td>$(HtmlEnc $r.NomFileShare)</td><td>$(HtmlEnc $r.TypeErreur)</td></tr>")
        }
        $footerAD = if ($DataAccessDenied.Count -gt 20) { "<tr class='table-footer'><td colspan='3'>... et $($DataAccessDenied.Count - 20) autres dossiers refuses.</td></tr>" } else { "" }
        $accessDeniedSection = @"
  <h3>&#x1F512; Dossiers refuses a l'audit ($($DataAccessDenied.Count))</h3>
  <p style="color:#c0392b;font-size:13px;">Ces dossiers n'ont pas pu etre audites. Leurs permissions NTFS sont inconnues. Decider du sort avant migration (P0).</p>
  <table>
    <tr><th>Chemin</th><th>FileShare</th><th>Type d'erreur</th></tr>
    $($adRows -join "`n    ")
    $footerAD
  </table>
"@
    }

    # Tableau des premières entrées ACL
    $tblRows = [System.Collections.Generic.List[string]]::new()
    foreach ($r in ($DataPermissions | Select-Object -First $MAX_ROWS)) {
        $tblRows.Add("<tr><td class='path-cell'>$(HtmlEnc $r.$cheminCol)</td><td>$(HtmlEnc $r.Proprietaire)</td><td>$(HtmlEnc $r.$idCol)</td><td>$(HtmlEnc $r.TypeAcces)</td><td>$(HtmlEnc $r.Droits)</td><td>$(HtmlEnc $r.EstHerite)</td></tr>")
    }
    $footer = if ($nb -gt $MAX_ROWS) { "<tr class='table-footer'><td colspan='6'>... et $($nb - $MAX_ROWS) autres entrees. Consultez : <code>$CsvName</code></td></tr>" } else { "" }
    return @"
<div class="section" id="section-permissions">
  <h2>&#x1F510; Permissions NTFS</h2>
  <div class="stat-row">
    <div class="stat-box"><div class="number">$nb</div><div class="label">Entrees ACL totales</div></div>
    <div class="stat-box ok"><div class="number">$pctHeritage%</div><div class="label">Heritage propre</div></div>
    <div class="stat-box warn"><div class="number">$nbRuptures</div><div class="label">Ruptures heritage</div></div>
    <div class="stat-box"><div class="number">$nbExplicite</div><div class="label">ACE explicites</div></div>
    <div class="stat-box"><div class="number">$nbIdentites</div><div class="label">Identites uniques</div></div>
  </div>
  $sidSection
  $accessDeniedSection
  $hotspotHtml
  $identitiesHtml
  <h3>Detail ACL (extrait, tri ruptures en premier)</h3>
  <p style="font-size:12px;color:#666;">$(if($CsvName){"Fichier complet : <code>$CsvName</code>"})</p>
  <table>
    <tr><th>Dossier</th><th>Proprietaire</th><th>Identite</th><th>Acces</th><th>Droits</th><th>Herite</th></tr>
    $($tblRows -join "`n    ")
    $footer
  </table>
</div>
"@
}

function Build-SectionHomeOwnership {
    param([array]$DataHomeOwnership, [string]$CsvName = "")
    if ($DataHomeOwnership.Count -eq 0) { return "" }

    $nbTotal     = $DataHomeOwnership.Count
    $nbActifs    = ($DataHomeOwnership | Where-Object { $_.Statut -like "Actif*" }       | Measure-Object).Count
    $nbDesact    = ($DataHomeOwnership | Where-Object { $_.Statut -like "D*sactiv*" }   | Measure-Object).Count
    $nbOrphelins = ($DataHomeOwnership | Where-Object { $_.Statut -eq "Orphelin" }        | Measure-Object).Count
    $nbService   = ($DataHomeOwnership | Where-Object { $_.Statut -eq "CompteService" }   | Measure-Object).Count
    $nbTechnique = ($DataHomeOwnership | Where-Object { $_.Statut -eq "DossierTechnique" }| Measure-Object).Count

    $statusGroups = $DataHomeOwnership | Group-Object Statut | Sort-Object Count -Descending
    $maxStat = if ($statusGroups) { ($statusGroups | Measure-Object Count -Maximum).Maximum } else { 1 }
    if ($maxStat -le 0) { $maxStat = 1 }

    $chartStatus = ""
    foreach ($g in $statusGroups) {
        $pct = [math]::Round(100.0 * $g.Count / $maxStat, 0)
        $cls = switch -Wildcard ($g.Name) {
            "Orphelin"         { "seg-error" }
            "D*sactiv*"        { "seg-warn" }
            "Actif_Inactif"    { "seg-warn" }
            "CompteService"    { "seg-warn" }
            default            { "seg-ok" }
        }
        $lbl = HtmlEnc $g.Name
        $chartStatus += "<div class='horiz-row'><div class='horiz-label'>$lbl</div><div class='horiz-track'><div class='horiz-fill $cls' style='width:${pct}%;'></div></div><div class='horiz-count'>$($g.Count)</div></div>"
    }

    $tblRows = [System.Collections.Generic.List[string]]::new()
    foreach ($r in ($DataHomeOwnership | Sort-Object Severite, NomDossier | Select-Object -First $MAX_ROWS)) {
        $badge = switch ($r.Severite) {
            "ERROR" { "<span class='badge badge-error'>ERROR</span>" }
            "WARN"  { "<span class='badge badge-warn'>WARN</span>" }
            default { "<span class='badge badge-ok'>OK</span>" }
        }
        $empl = Get-EmplacementCell $r.CheminComplet
        $tblRows.Add("<tr><td>$(HtmlEnc $r.NomDossier)</td><td>$(HtmlEnc $r.Statut)</td><td>$(HtmlEnc $r.DisplayName)</td><td>$(HtmlEnc $r.Email)</td><td>$(HtmlEnc $r.Departement)</td><td>$(HtmlEnc $r.Enabled)</td><td>$(HtmlEnc $r.LastLogonDate)</td><td>$($r.TailleGB)</td><td>$badge</td><td>$(HtmlEnc $r.Recommandation)</td>$empl</tr>")
    }
    $footer = if ($nbTotal -gt $MAX_ROWS) { "<tr class='table-footer'><td colspan='11'>... et $($nbTotal - $MAX_ROWS) autres elements. Consultez : <code>$CsvName</code></td></tr>" } else { "" }

    return @"
<div class="section" id="section-homes-ownership">
  <h2>&#x1F464; Dossiers Personnels &mdash; Verification AD</h2>
  <div class="stat-row">
    <div class="stat-box"><div class="number">$nbTotal</div><div class="label">Homes analyses</div></div>
    <div class="stat-box ok"><div class="number">$nbActifs</div><div class="label">Comptes actifs</div></div>
    <div class="stat-box warn"><div class="number">$nbDesact</div><div class="label">Comptes desactives</div></div>
    <div class="stat-box danger"><div class="number">$nbOrphelins</div><div class="label">Orphelins</div></div>
    <div class="stat-box warn"><div class="number">$nbService</div><div class="label">Comptes service</div></div>
    <div class="stat-box"><div class="number">$nbTechnique</div><div class="label">Dossiers techniques</div></div>
  </div>
  <h3>Repartition des statuts</h3>
  <div class="chart-horiz">$chartStatus</div>
  <h3>Detail ($nbTotal homes)</h3>
  <div class="section-counter"><span class='badge badge-error'>$nbOrphelins orphelins</span>&nbsp;<span class='badge badge-warn'>$($nbDesact + $nbService) a traiter</span>&nbsp;<span class='badge badge-ok'>$nbActifs actifs</span></div>
  <table>
    <tr><th>Dossier</th><th>Statut</th><th>DisplayName</th><th>Email</th><th>Departement</th><th>Active</th><th>Derniere connexion</th><th>Taille (GB)</th><th>Severite</th><th>Recommandation</th><th>Emplacement</th></tr>
    $($tblRows -join "`n    ")
    $footer
  </table>
</div>
"@
}

function Build-SectionHomeStats {
    param([array]$DataHomeStats, [string]$CsvName = "")
    if ($DataHomeStats.Count -eq 0) { return "" }

    $nbTotal    = $DataHomeStats.Count
    $catGroups  = $DataHomeStats | Group-Object Categorie | Sort-Object Count -Descending
    $maxCat = if ($catGroups) { ($catGroups | Measure-Object Count -Maximum).Maximum } else { 1 }
    if ($maxCat -le 0) { $maxCat = 1 }

    $statRowCat = ""; $chartCat = ""
    foreach ($g in $catGroups) {
        $pct = [math]::Round(100.0 * $g.Count / $maxCat, 0)
        $cls = switch -Wildcard ($g.Name) {
            "Abandon*" { "seg-error" }
            "Dormant*" { "seg-warn" }
            "Semi*"    { "seg-warn" }
            default    { "seg-ok" }
        }
        $emoji = ($g.Group | Select-Object -First 1).CategorieEmoji
        if (-not $emoji) { $emoji = "" }
        $volGo = [math]::Round(($g.Group | ForEach-Object { [double](($_.TailleGB -replace ',','.') -replace '[^0-9.]','0') } | Measure-Object -Sum).Sum, 1)
        $lbl = "$emoji $(HtmlEnc $g.Name)"
        $statRowCat += "<div class='stat-box'><div class='number'>$($g.Count)</div><div class='label'>$lbl ($volGo GB)</div></div>"
        $chartCat   += "<div class='horiz-row'><div class='horiz-label'>$lbl</div><div class='horiz-track'><div class='horiz-fill $cls' style='width:${pct}%;'></div></div><div class='horiz-count'>$($g.Count)</div></div>"
    }

    $tblRows = [System.Collections.Generic.List[string]]::new()
    foreach ($r in ($DataHomeStats | Where-Object { $_.TailleGB -match '[\d,\.]' } | Sort-Object { [double](($_.TailleGB -replace ',','.') -replace '[^0-9.]','0') } -Descending | Select-Object -First $MAX_ROWS)) {
        $empl  = Get-EmplacementCell $r.CheminComplet
        $emoji = if ($r.CategorieEmoji) { $r.CategorieEmoji } else { "" }
        $tblRows.Add("<tr><td>$(HtmlEnc $r.NomDossier)</td><td>$emoji $(HtmlEnc $r.Categorie)</td><td>$($r.TailleGB)</td><td>$(HtmlEnc $r.Recommandation)</td>$empl</tr>")
    }
    $footer = if ($DataHomeStats.Count -gt $MAX_ROWS) { "<tr class='table-footer'><td colspan='5'>... et $($DataHomeStats.Count - $MAX_ROWS) autres homes. Consultez : <code>$CsvName</code></td></tr>" } else { "" }

    return @"
<div class="section" id="section-homes-stats">
  <h2>&#x1F4CA; Dossiers Personnels &mdash; Activite</h2>
  <div class="stat-row">$statRowCat</div>
  <h3>Repartition par categorie d'activite</h3>
  <div class="chart-horiz">$chartCat</div>
  <h3>Top $MAX_ROWS homes les plus volumineux</h3>
  <table>
    <tr><th>Dossier</th><th>Categorie</th><th>Taille (GB)</th><th>Recommandation</th><th>Emplacement</th></tr>
    $($tblRows -join "`n    ")
    $footer
  </table>
</div>
"@
}

function Build-SectionAgeFichiers {
    param(
        [string]$AgeSummaryPath = "",
        [array]$AgeDetails      = @(),
        [string]$NomFS          = ""
    )

    # Try to load summary
    $summaryData = @()
    if ($AgeSummaryPath -and (Test-Path $AgeSummaryPath)) {
        try { $summaryData = @(Import-Csv -Path $AgeSummaryPath -Delimiter ";" -Encoding UTF8) } catch {}
    }

    if ($summaryData.Count -eq 0 -and $AgeDetails.Count -eq 0) { return "" }

    # Helper to get a column value safely
    function _col { param($row, $names)
        foreach ($n in $names) { if ($row.PSObject.Properties[$n]) { return $row.$n } }
        return ""
    }

    # Stat boxes for 2/5/10 ans from summary
    $statBoxes = ""
    if ($summaryData.Count -gt 0) {
        foreach ($seuil in @("2ans","5ans","10ans")) {
            $row = $summaryData | Where-Object { (_col $_ @("Seuil","seuil","Periode","periode")) -like "*$seuil*" } | Select-Object -First 1
            if (-not $row) { $row = $summaryData | Where-Object { (_col $_ @("NonModifieDepuis","non_modifie_depuis")) -like "*$seuil*" } | Select-Object -First 1 }
            if ($row) {
                $nbF  = _col $row @("NombreFichiers","nombre_fichiers","Fichiers","NbFichiers")
                $vol  = _col $row @("TailleGB","taille_gb","VolumeGB","Volume","TailleLisible")
                $cls  = if ($seuil -eq "10ans") { "danger" } elseif ($seuil -eq "5ans") { "warn" } else { "" }
                $statBoxes += "<div class='stat-box $cls'><div class='number'>$(HtmlEnc $nbF)</div><div class='label'>Non modifies &gt; $seuil ($vol GB)</div></div>"
            }
        }
    }

    # Top 10 N1 avec le plus de fichiers > 5 ans (depuis AgeDetails)
    $top10Html = ""
    $top10Data = @{}
    foreach ($f in ($AgeDetails | Where-Object { $_.Name -like "*5ans*" })) {
        try {
            $rows = @(Import-Csv -Path $f.FullName -Delimiter ";" -Encoding UTF8)
            foreach ($r in $rows) {
                $dossier = _col $r @("DossierNiveau1","Dossier","N1","dossier_n1")
                if ($dossier) {
                    if (-not $top10Data.ContainsKey($dossier)) { $top10Data[$dossier] = 0 }
                    $top10Data[$dossier]++
                }
            }
        } catch {}
    }

    if ($top10Data.Count -gt 0) {
        $top10Sorted = $top10Data.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 10
        $maxVal = ($top10Sorted | Measure-Object Value -Maximum).Maximum
        if ($maxVal -le 0) { $maxVal = 1 }
        $top10Html = "<h3>Top 10 dossiers N1 par nb fichiers non modifies &gt; 5 ans</h3><div class='chart-horiz'>"
        foreach ($entry in $top10Sorted) {
            $pct = [math]::Round(100.0 * $entry.Value / $maxVal, 0)
            $lbl = HtmlEnc $entry.Key
            $top10Html += "<div class='horiz-row'><div class='horiz-label'>$lbl</div><div class='horiz-track'><div class='horiz-fill seg-warn' style='width:${pct}%;'></div></div><div class='horiz-count'>$($entry.Value)</div></div>"
        }
        $top10Html += "</div>"
    }

    $infoMsg = "<p style='font-size:12px;color:#666;margin-top:12px;'>&#x2139; Un fichier non modifie depuis 12 ans apparait simultanement dans les CSV 2 ans, 5 ans et 10 ans (cumul). Les totaux peuvent donc etre superieurs au nombre reel de fichiers distincts.</p>"

    # If no stat boxes and no chart, show a friendly empty-state message
    $noDataMsg = ""
    if (-not $statBoxes -and -not $top10Html) {
        $noDataMsg = "<p style='color:#999;font-style:italic;'>Aucun fichier ancien detecte (ou donnees d''age non disponibles pour ce FileShare).</p>"
    }

    return @"
<div class="section" id="section-age-fichiers">
  <h2>&#x1F4C5; Age des Fichiers (non modifies)</h2>
  <div class="stat-row">$statBoxes</div>
  $noDataMsg
  $top10Html
  $infoMsg
</div>
"@
}

function Build-SectionPyramideAges {
    param(
        [string]$InventaireDetailPath = "",
        [string]$NomFS = ""
    )

    $fs = if ($NomFS) { HtmlEnc $NomFS } else { "" }
    $missingDataMessage = @"
<div class="section" id="section-pyramide-ages">
  <h2>&#x1F4C5; Pyramide des ages des fichiers$(if ($fs) {" &mdash; $fs"} else {""})</h2>
  <div class="info-block" style="background:#fff6e6;border-left:4px solid #e67e22;padding:12px 16px;margin:8px 0 16px;border-radius:4px;font-size:13px;">
    Aucune donnee detaillee disponible &mdash; relancer <code>Get-FileShareInventory.ps1 -IncludeFileDetail</code> pour generer le CSV <code>Inventaire_FileShare_Detail_&lt;NomFS&gt;_*.csv</code>.
  </div>
</div>
"@

    if ([string]::IsNullOrWhiteSpace($InventaireDetailPath) -or -not (Test-Path $InventaireDetailPath)) {
        Write-Log "[Pyramide] $NomFS : CSV detail introuvable" "WARN"
        return $missingDataMessage
    }

    try {
        $rows = @(Import-Csv -Path $InventaireDetailPath -Delimiter ';' -Encoding UTF8)
    } catch {
        Write-Log "[Pyramide] $NomFS : lecture CSV detail impossible ($InventaireDetailPath) : $_" "WARN"
        return $missingDataMessage
    }

    if ($rows.Count -eq 0) {
        Write-Log "[Pyramide] $NomFS : CSV detail vide ($InventaireDetailPath)" "WARN"
        return $missingDataMessage
    }

    $now = Get-Date
    $buckets = [ordered]@{
        "moins1an"  = @{ Label = "&lt;1 an";   NbFichiers = 0; TailleOctets = [long]0 }
        "1a2ans"    = @{ Label = "1-2 ans";   NbFichiers = 0; TailleOctets = [long]0 }
        "2a5ans"    = @{ Label = "2-5 ans";   NbFichiers = 0; TailleOctets = [long]0 }
        "5a10ans"   = @{ Label = "5-10 ans";  NbFichiers = 0; TailleOctets = [long]0 }
        "plus10ans" = @{ Label = "&gt;10 ans"; NbFichiers = 0; TailleOctets = [long]0 }
    }

    $nbLignesLues = 0
    foreach ($row in $rows) {
        $nbLignesLues++
        $dateVal = "$($row.DateModification)"
        if ([string]::IsNullOrWhiteSpace($dateVal)) { continue }

        $dateObj = $null
        try {
            $dateObj = [datetime]::ParseExact($dateVal.Trim(), 'yyyy-MM-dd HH:mm:ss', [System.Globalization.CultureInfo]::InvariantCulture)
        } catch {
            $fallbackDt = [datetime]::MinValue
            if (-not [datetime]::TryParse($dateVal, [ref]$fallbackDt)) { continue }
            $dateObj = $fallbackDt
        }

        $tailleNum = ConvertTo-LongSafe "$($row.TailleOctets)"
        $ageYears = ($now - $dateObj).TotalDays / 365.25

        $bucket = if ($ageYears -lt 1) { "moins1an" }
                  elseif ($ageYears -lt 2) { "1a2ans" }
                  elseif ($ageYears -lt 5) { "2a5ans" }
                  elseif ($ageYears -lt 10) { "5a10ans" }
                  else { "plus10ans" }

        $buckets[$bucket].NbFichiers++
        $buckets[$bucket].TailleOctets += $tailleNum
    }

    $totalCount = [long](($buckets.Values | Measure-Object -Property NbFichiers -Sum).Sum)
    $totalBytes = [long](($buckets.Values | Measure-Object -Property TailleOctets -Sum).Sum)

    if ($totalCount -eq 0) {
        Write-Log "[Pyramide] $NomFS : $nbLignesLues lignes lues, aucun bucket exploitable" "WARN"
        return $missingDataMessage
    }

    $migrableCount = $buckets["moins1an"].NbFichiers + $buckets["1a2ans"].NbFichiers
    $migrableBytes = $buckets["moins1an"].TailleOctets + $buckets["1a2ans"].TailleOctets
    $horsScopeCount = $buckets["2a5ans"].NbFichiers + $buckets["5a10ans"].NbFichiers + $buckets["plus10ans"].NbFichiers
    $horsScopeBytes = $buckets["2a5ans"].TailleOctets + $buckets["5a10ans"].TailleOctets + $buckets["plus10ans"].TailleOctets

    $pctMigrable = [math]::Round((100.0 * $migrableCount / $totalCount), 1)
    $pctHorsScope = [math]::Round((100.0 * $horsScopeCount / $totalCount), 1)

    $maxBucket = ($buckets.Values | Measure-Object -Property NbFichiers -Maximum).Maximum
    if (-not $maxBucket -or $maxBucket -le 0) { $maxBucket = 1 }

    $chartRows = [System.Collections.Generic.List[string]]::new()
    foreach ($bk in $buckets.Keys) {
        $b = $buckets[$bk]
        $pct = [math]::Round(100.0 * $b.NbFichiers / $maxBucket, 0)
        $cls = if ($bk -in @("moins1an","1a2ans")) { "seg-ok" } elseif ($bk -eq "plus10ans") { "seg-error" } else { "seg-warn" }
        $vol = Format-VolumeReadable ($b.TailleOctets / 1073741824.0)
        $chartRows.Add("<div class='horiz-row'><div class='horiz-label'>$($b.Label)</div><div class='horiz-track'><div class='horiz-fill $cls' style='width:${pct}%;'></div></div><div class='horiz-count'>$($b.NbFichiers) <span style='color:#999;font-size:10px;'>($vol)</span></div></div>")
    }
    $chartHtml = $chartRows -join ""

    $tableRows = [System.Collections.Generic.List[string]]::new()
    foreach ($bk in $buckets.Keys) {
        $b = $buckets[$bk]
        $pct = [math]::Round((100.0 * $b.NbFichiers / $totalCount), 1)
        $scope = if ($bk -in @("moins1an","1a2ans")) { "<span class='badge badge-ok'>A migrer</span>" } else { "<span class='badge badge-warn'>Hors perimetre</span>" }
        $tableRows.Add("<tr><td>$($b.Label)</td><td>$($b.NbFichiers)</td><td>${pct}%</td><td>$(Format-VolumeReadable ($b.TailleOctets / 1073741824.0))</td><td>$scope</td></tr>")
    }
    $tableRows.Add("<tr class='tbl-total'><td><strong>[TOTAL]</strong></td><td><strong>$totalCount</strong></td><td><strong>100%</strong></td><td><strong>$(Format-VolumeReadable ($totalBytes / 1073741824.0))</strong></td><td><strong>-</strong></td></tr>")
    $tableRowsHtml = $tableRows -join "`n    "

    Write-Log ("[Pyramide] {0} : {1} lignes lues, buckets : <1an={2}, 1-2ans={3}, 2-5ans={4}, 5-10ans={5}, >10ans={6}" -f `
        $NomFS, $nbLignesLues, $buckets["moins1an"].NbFichiers, $buckets["1a2ans"].NbFichiers, $buckets["2a5ans"].NbFichiers, $buckets["5a10ans"].NbFichiers, $buckets["plus10ans"].NbFichiers) "INFO"

    return @"
<div class="section" id="section-pyramide-ages">
  <h2>&#x1F4C5; Pyramide des ages des fichiers$(if ($fs) {" &mdash; $fs"} else {""})</h2>
  <div class="info-block" style="background:#eafaf1;border-left:4px solid #27ae60;padding:12px 16px;margin:8px 0 16px;border-radius:4px;font-size:13px;">
    <strong>&#x2139; Decision client :</strong> Seuls les fichiers modifies depuis moins de 2 ans seront migres vers M365.
    Les fichiers plus anciens resteront sur le FileShare (archivage) ou seront purges selon arbitrage.
  </div>
  <div class="stat-row">
    <div class="stat-box ok">
      <div class="number">$migrableCount</div>
      <div class="label">A MIGRER (&lt;2 ANS) &mdash; $migrableCount fichiers ($(Format-VolumeReadable ($migrableBytes / 1073741824.0))) &mdash; $($pctMigrable)%</div>
    </div>
    <div class="stat-box warn">
      <div class="number">$horsScopeCount</div>
      <div class="label">HORS PERIMETRE (&ge;2 ANS) &mdash; $horsScopeCount fichiers ($(Format-VolumeReadable ($horsScopeBytes / 1073741824.0))) &mdash; $($pctHorsScope)%</div>
    </div>
    <div class="stat-box">
      <div class="number">$totalCount</div>
      <div class="label">TOTAL FICHIERS ANALYSES &mdash; $totalCount fichiers ($(Format-VolumeReadable ($totalBytes / 1073741824.0)))</div>
    </div>
  </div>
  <h3>Repartition par tranche d'age (date de derniere modification)</h3>
  <div class="chart-horiz">$chartHtml</div>
  <table style="margin-top:12px;">
    <tr><th>Tranche d'age</th><th>Nb fichiers</th><th>% du total</th><th>Volume</th><th>Perimetre migration</th></tr>
    $tableRowsHtml
  </table>
</div>
"@
}

function Build-Navigation {
    param(
        [bool]$HasInventaire  = $false,
        [bool]$HasChemins     = $false,
        [bool]$HasDoublons    = $false,
        [bool]$HasExtensions  = $false,  # conserve pour compatibilite mais ignore en v3.3
        [bool]$HasPermissions = $false,
        [bool]$HasHomes       = $false,
        [bool]$HasHomesStats  = $false,
        [bool]$HasAgeFichiers = $false,
        [bool]$HasPyramideAges = $false
    )
    $links = [System.Collections.Generic.List[string]]::new()
    $links.Add('<a href="#section-readiness">&#x1F3AF; Readiness</a>')
    $links.Add('<a href="#section-summary">&#x1F4DD; Synthese</a>')
    $links.Add('<a href="#section-charts">&#x1F4CA; Graphiques</a>')
    $links.Add('<a href="#section-plan">&#x1F3AF; Plan d''action</a>')
    if ($HasInventaire)   { $links.Add('<a href="#section-inventaire">&#x1F4E6; Inventaire</a>') }
    if ($HasChemins)      { $links.Add('<a href="#section-chemins">&#x1F4CF; Chemins Trop Longs</a>') }
    if ($HasDoublons)     { $links.Add('<a href="#section-doublons">&#x1F4D1; Doublons</a>') }
    # v3.3 — Extensions a auditer ne sont plus affichees dans la nav
    if ($HasPermissions)  { $links.Add('<a href="#section-permissions">&#x1F510; Permissions NTFS</a>') }
    if ($HasHomes)        { $links.Add('<a href="#section-homes-ownership">&#x1F464; Homes (AD)</a>') }
    if ($HasHomesStats)   { $links.Add('<a href="#section-homes-stats">&#x1F4CA; Homes (Activite)</a>') }
    if ($HasAgeFichiers)  { $links.Add('<a href="#section-age-fichiers">&#x1F4C5; Age Fichiers</a>') }
    if ($HasPyramideAges) { $links.Add('<a href="#section-pyramide-ages">&#x1F4C5; Pyramide des ages</a>') }
    return "<nav id='nav-sticky'>" + ($links -join " ") + "</nav>"
}

function Build-Footer {
    param([string]$CheminSource = "")
    $now = Get-Date -Format "dd/MM/yyyy HH:mm:ss"
    $srcInfo = if ($CheminSource) { "<br>Source : <code>$(HtmlEnc $CheminSource)</code>" } else { "" }
    return @"
<div class="report-footer">
  <p>Rapport genere le <strong>$now</strong>$srcInfo</p>
  <p>Version : <strong>v3.4 &mdash; Dossier rapports configurable + permissions NTFS systematiques + filtrage UNC robuste</strong></p>
  <p>PrimaGAZ &times; Avanade &mdash; Propulse par ShareGate</p>
</div>
"@
}

# ============================================================
# ASSEMBLAGE DU RAPPORT COMPLET
# ============================================================

function Build-FullReport {
    param(
        [string]$Title          = "Dashboard Assessment — PrimaGAZ",
        [string]$Subtitle       = "Migration FileShare vers M365",
        [array]$Inventaire      = @(),
        [array]$Volumineux      = @(),
        [array]$Invalides       = @(),
        [array]$Chemins         = @(),
        [array]$Doublons        = @(),
        [array]$Extensions      = @(),
        [array]$Permissions     = @(),
        [array]$SidResolution   = @(),
        [array]$AccessDenied    = @(),
        [array]$HomeOwnership   = @(),
        [array]$HomeStats       = @(),
        [array]$AgeFichiers     = @(),
        [array]$AgeDetails      = @(),
        [string]$AgeSummaryPath = "",
        [string]$InventaireDetailPath = "",
        [string]$CheminUNCRoot  = "",
        [string]$CsvVolName     = "",
        [string]$CsvInvName     = "",
        [string]$CsvInvalName   = "",
        [string]$CsvChmName     = "",
        [string]$CsvDblName     = "",
        [string]$CsvExtName     = "",
        [string]$CsvPrmName     = "",
        [string]$CsvHOwName     = "",
        [string]$CsvHStName     = "",
        [string]$CheminSource   = "",
        [string]$NomFS          = "",
        [string]$OutputPath     = "",
        [bool]$HasPermissionsSource = $false
    )

    $ligneTotal    = $Inventaire | Where-Object {
        $_.DossierNiveau1 -eq "[TOTAL]" -and (
            -not $_.PSObject.Properties['NomFileShare'] -or $_.NomFileShare -eq "[CONSOLIDE]"
        )
    } | Select-Object -First 1
    if (-not $ligneTotal) {
        $ligneTotal = $Inventaire | Where-Object { $_.DossierNiveau1 -eq "[TOTAL]" } | Select-Object -First 1
    }
    $totalFichiers = if ($ligneTotal -and $ligneTotal.NombreFichiers -match '^\d+$') { [long]$ligneTotal.NombreFichiers } else { 0 }
    $volumeOctets  = if ($ligneTotal) { ConvertTo-LongSafe $ligneTotal.TailleOctets } else { 0 }
    $volumeGB      = [math]::Round($volumeOctets / 1073741824.0, 2)

    $nbExtensions    = $Extensions.Count
    # v3.3 — Extensions ne comptent plus dans les warnings (affichage supprime)
    $nbInvalidesError= ($Invalides | Where-Object { $_.Severite -eq "ERROR" } | Measure-Object).Count
    $nbInvalidesWarn = ($Invalides | Where-Object { $_.Severite -eq "WARN"  } | Measure-Object).Count
    # Items INFO (FichierTemporaire, FichierSysteme, FichierVerrouillage) exclus des stats bloquantes
    $nbChemins       = $Chemins.Count
    $nbVolumineux    = $Volumineux.Count
    $nbVolBloquant   = ($Volumineux | Where-Object { $_.Severite -eq "ERROR" } | Measure-Object).Count
    $nbDoublons      = $Doublons.Count
    $nbPrefixes      = ($Invalides | Where-Object { $_.ProblemeType -eq "PrefixeInterdit" } | Measure-Object).Count
    $nbFichSys       = ($Invalides | Where-Object { $_.ProblemeType -in @("FichierSysteme","FichierTemporaire","FichierVerrouillage") } | Measure-Object).Count

    # Recalcul readiness : seules les vraies ERROR comptent (pas les extensions ni les INFO)
    $blockingErrors  = $nbInvalidesError + $nbChemins
    # v3.3 — Extensions ne contribuent plus aux warnings
    $warnings        = $nbInvalidesWarn + $nbVolumineux + ($Doublons | Select-Object -ExpandProperty GroupeDoublonId -Unique | Measure-Object).Count

    # Ruptures d'heritage NTFS → warnings
    $cheminColPrm = if ($Permissions.Count -gt 0 -and $Permissions[0].PSObject.Properties.Name -contains "CheminDossier") { "CheminDossier" } else { "Path" }
    $nbRupturesHeritage = if ($Permissions.Count -gt 0) {
        ($Permissions | Where-Object { $_.EstHerite -ne "True" -and $_.EstHerite -ne "1" } |
            Select-Object -ExpandProperty $cheminColPrm -Unique | Where-Object { $_ } | Measure-Object).Count
    } else { 0 }
    $warnings += $nbRupturesHeritage

    $nbOrphelins   = ($HomeOwnership | Where-Object { $_.Statut -eq "Orphelin" }     | Measure-Object).Count
    $goOrphelins   = [math]::Round(($HomeOwnership | Where-Object { $_.Statut -eq "Orphelin" } | ForEach-Object { [double](($_.TailleGB -replace ',','.') -replace '[^0-9.]','0') } | Measure-Object -Sum).Sum, 1)
    $nbDesactives  = ($HomeOwnership | Where-Object { $_.Statut -like "D*sactiv*" }    | Measure-Object).Count
    $nbInactifsHome= ($HomeOwnership | Where-Object { $_.Statut -eq "Actif_Inactif" } | Measure-Object).Count
    $nbAbandonnes  = ($HomeStats     | Where-Object { $_.Categorie -like "Abandon*" } | Measure-Object).Count
    $goAbandonnes  = [math]::Round(($HomeStats | Where-Object { $_.Categorie -like "Abandon*" } | ForEach-Object { [double](($_.TailleGB -replace ',','.') -replace '[^0-9.]','0') } | Measure-Object -Sum).Sum, 1)
    $nbVidesHome   = ($HomeStats     | Where-Object { $_.Categorie -eq "Vide" }      | Measure-Object).Count

    if ($HomeOwnership.Count -gt 0 -or $HomeStats.Count -gt 0) {
        $blockingErrors += $nbOrphelins
        $warnings       += $nbDesactives + $nbInactifsHome
    }

    # AccessDenied et SID inconnus
    $nbAccessDenied = $AccessDenied.Count
    $nbSidInconnus  = ($SidResolution | Where-Object { $_.Statut -eq "DOMAINE_INCONNU" } | Measure-Object).Count
    if ($nbAccessDenied -gt 0) { $blockingErrors += 1 }  # Au moins 1 dossier inaccessible = blocker
    if ($nbSidInconnus  -gt 0) { $blockingErrors += $nbSidInconnus }

    $totalItems = [math]::Max($totalFichiers, $blockingErrors + $warnings + 1)

    $hasPermissionsSection = $HasPermissionsSource -or ($Permissions.Count -gt 0)

    $nav = Build-Navigation `
        -HasInventaire  ($Inventaire.Count  -gt 0) `
        -HasChemins     ($Chemins.Count     -gt 0) `
        -HasDoublons    ($Doublons.Count    -gt 0) `
        -HasExtensions  $false `
        -HasPermissions $hasPermissionsSection `
        -HasHomes       ($HomeOwnership.Count -gt 0) `
        -HasHomesStats  ($HomeStats.Count   -gt 0) `
        -HasAgeFichiers ($AgeSummaryPath -ne "" -or $AgeDetails.Count -gt 0) `
        -HasPyramideAges ($NomFS -ne "" -or (-not [string]::IsNullOrWhiteSpace($InventaireDetailPath) -and (Test-Path $InventaireDetailPath)))

    $summaryNomFs = if ([string]::IsNullOrWhiteSpace($NomFS)) { $Subtitle } else { $NomFS }
    $sReadiness  = Build-SectionReadiness -BlockingErrors $blockingErrors -Warnings $warnings -TotalItems $totalItems -VolumeGB $volumeGB -Scope $Subtitle
    $sSummary    = Build-SectionSummary   -TotalFichiers $totalFichiers -VolumeGB $volumeGB -BlockingErrors $blockingErrors -Warnings $warnings -NbExtensions $nbExtensions -NbInvalides ($Invalides.Count) -NbChemins $nbChemins -NbVolumineux $nbVolumineux -NbDoublons $nbDoublons -NbRupturesHeritage $nbRupturesHeritage -Scope $Subtitle -NomFS $summaryNomFs
    $sCharts     = Build-SectionCharts    -DataInvalides $Invalides -DataExtensions $Extensions -DataChemins $Chemins -DataVolumineux $Volumineux -DataInventaire $Inventaire -NbRupturesHeritage $nbRupturesHeritage
    $sPlan       = Build-SectionActionPlan -NbExtensions $nbExtensions -NbInvalidesError $nbInvalidesError -NbCheminsError $nbChemins -NbVolumineuxBloquant $nbVolBloquant -NbDoublons $nbDoublons -NbPrefixesInterdits $nbPrefixes -NbFichiersSysteme $nbFichSys -NbOrphelins $nbOrphelins -GoOrphelins $goOrphelins -NbDesactives $nbDesactives -NbInactifsHome $nbInactifsHome -NbAbandonnes $nbAbandonnes -GoAbandonnes $goAbandonnes -NbVides $nbVidesHome -NbAccessDenied $nbAccessDenied -NbSidInconnus $nbSidInconnus -HasPermissions $hasPermissionsSection -NbRupturesHeritage $nbRupturesHeritage

    $sInventaire = Build-SectionInventaire  -DataInventaire $Inventaire
    $sVolumineux = Build-SectionVolumineux  -DataVolumineux $Volumineux  -CsvName $CsvVolName
    $sInvalides  = Build-SectionInvalides   -DataInvalides  $Invalides   -CsvName $CsvInvalName
    $sChemins    = Build-SectionChemins     -DataChemins    $Chemins     -CsvName $CsvChmName
    $sDoublons   = Build-SectionDoublons    -DataDoublons   $Doublons    -CsvName $CsvDblName
    $sExtensions = Build-SectionExtensions  -DataExtensions $Extensions  -CsvName $CsvExtName  # v3.3 : construit mais pas affiché
    $sPerms      = Build-SectionPermissions -DataPermissions $Permissions -CsvName $CsvPrmName -DataSidResolution $SidResolution -DataAccessDenied $AccessDenied -NomFS $NomFS -OutputPath $OutputPath -Timestamp $timestamp -CheminUNCRoot $CheminUNCRoot -PermissionsCsvAvailable $hasPermissionsSection
    $sHomeOwnrsh = Build-SectionHomeOwnership -DataHomeOwnership $HomeOwnership -CsvName $CsvHOwName
    $sHomeStats   = Build-SectionHomeStats   -DataHomeStats  $HomeStats   -CsvName $CsvHStName
    $sAgeFichiers = Build-SectionAgeFichiers -AgeSummaryPath $AgeSummaryPath -AgeDetails $AgeDetails
    $sPyramide    = Build-SectionPyramideAges -InventaireDetailPath $InventaireDetailPath -NomFS $NomFS
    $sFooter      = Build-Footer             -CheminSource $CheminSource

    $css = Get-CommonCss
    $now = Get-Date -Format "dd/MM/yyyy HH:mm"
    return @"
<!DOCTYPE html>
<html lang="fr">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>$(HtmlEnc $Title)</title>
  <style>$css</style>
</head>
<body>
$nav
<div class="page-header">
  <h1>$(HtmlEnc $Title)</h1>
  <p>$(HtmlEnc $Subtitle) | Genere le $now | Outil : ShareGate</p>
</div>
$sReadiness
$sSummary
$sCharts
$sPlan
$sInventaire
$sVolumineux
$sInvalides
$sChemins
$sDoublons
$sPerms
$sHomeOwnrsh
$sHomeStats
$sAgeFichiers
$sPyramide
$sFooter
</body>
</html>
"@
}

# ============================================================
# FONCTION : INDEX DSI (mode split)
# ============================================================

function Build-IndexReport {
    param(
        [array]$Groups,
        [string]$GlobalFile   = "",
        [array]$FsMappings    = @(),
        [array]$UnmappedPaths = @()
    )
    $css = Get-CommonCss
    $totalBlocking = ([long]($Groups | Measure-Object -Property BlockingErrors -Sum).Sum)
    $totalItems    = ([long]($Groups | Measure-Object -Property TotalItems     -Sum).Sum)
    $totalVolumeGB = [math]::Round(($Groups | Measure-Object -Property VolumeGB -Sum).Sum, 1)
    $rGlobal = Get-ReadinessInfo -BlockingErrors $totalBlocking -TotalItems $totalItems

    $rowsByFs = [System.Collections.Generic.List[string]]::new()
    foreach ($g in ($Groups | Sort-Object NomFS)) {
        $r = Get-ReadinessInfo -BlockingErrors $g.BlockingErrors -TotalItems $g.TotalItems
        $volStr = Format-VolumeReadable $g.VolumeGB
        $reportLink = if ($g.ReportFile) { "<a href='$(HtmlEnc $g.ReportFile)'>&#x1F4C4; Voir</a>" } else { "&mdash;" }
        $readiness  = "$($r.Icon) <span style='color:$($r.Color);font-weight:600;'>$($r.Pct)%</span>"
        $rowsByFs.Add("<tr><td><strong>$(HtmlEnc $g.NomFS)</strong></td><td>$(HtmlEnc $g.TypeUsage)</td><td>$(HtmlEnc $g.CibleM365)</td><td>$volStr</td><td><span class='badge badge-error'>$($g.BlockingErrors)</span></td><td><span class='badge badge-warn'>$($g.Warnings)</span></td><td>$readiness</td><td>$reportLink</td></tr>")
    }

    $unmappedSection = ""
    $uniqueUnmapped = @($UnmappedPaths | Select-Object -Unique)
    if ($uniqueUnmapped.Count -gt 0) {
        $unmappedRows = ($uniqueUnmapped | Select-Object -First 50 | ForEach-Object { "<li><code>$(HtmlEnc $_)</code></li>" }) -join ""
        $unmappedSection = @"
<div class="section">
  <h2>&#x26A0; Chemins non couverts par le mapping</h2>
  <div class="unmapped-alert">
    <strong>$($uniqueUnmapped.Count) chemins uniques</strong> ne sont couverts par aucune entree du FileShareMapping.csv.
    Completez le fichier <code>FileShareMapping.csv</code> pour les integrer dans les rapports par FileShare.
  </div>
  <ul style="margin:12px 0 0 20px;font-size:13px;line-height:2;">$unmappedRows</ul>
</div>
"@
    }

    $globalLink = if ($GlobalFile) { "<a href='$(HtmlEnc $GlobalFile)' style='color:white;font-size:13px;'>&#x1F4C4; Rapport global confidentiel</a>" } else { "" }
    $now = Get-Date -Format "dd/MM/yyyy HH:mm"
    $nowFull = Get-Date -Format "dd/MM/yyyy HH:mm:ss"

    return @"
<!DOCTYPE html>
<html lang="fr">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Index Assessment &mdash; PrimaGAZ DSI</title>
  <style>$css</style>
</head>
<body>
<div class="page-header">
  <h1>&#x1F5C2; Index Assessment &mdash; PrimaGAZ DSI</h1>
  <p>Vue comparative multi-FileShares | Genere le $now | $globalLink</p>
</div>
<div class="section" id="section-readiness">
  <h2>&#x1F3AF; Readiness Globale</h2>
  <div class="readiness-block" style="background:$($rGlobal.BgColor);border:2px solid $($rGlobal.Color);">
    <div class="readiness-icon">$($rGlobal.Icon)</div>
    <div>
      <div class="readiness-pct" style="color:$($rGlobal.Color);">$($rGlobal.Pct)%</div>
      <h3 style="color:$($rGlobal.Color);font-size:20px;margin:4px 0;">$(HtmlEnc $rGlobal.Text)</h3>
      <p>Perimetre complet &mdash; $totalItems items, $totalBlocking erreurs bloquantes, $(Format-VolumeReadable $totalVolumeGB)</p>
    </div>
  </div>
</div>
<div class="section">
  <h2>&#x1F4C1; Tableau comparatif &mdash; FileShares ($($Groups.Count))</h2>
  <table>
    <tr><th>NomFS</th><th>Type usage</th><th>Cible M365</th><th>Volume</th><th>Erreurs</th><th>Warnings</th><th>Readiness</th><th>Rapport</th></tr>
    $($rowsByFs -join "`n    ")
  </table>
</div>
$unmappedSection
<div class="report-footer">
  <p>PrimaGAZ &times; Avanade &mdash; Propulse par ShareGate | v3.4 &mdash; ReportOutputPath + permissions NTFS systematiques</p>
  <p>Genere le $nowFull</p>
</div>
</body>
</html>
"@
}

# ============================================================
# LOGIQUE PRINCIPALE
# ============================================================

$csvVolName  = if ($csvVolumineux)    { $csvVolumineux.Name }    else { "" }
$csvInvName  = if ($csvInventaire)    { $csvInventaire.Name }    else { "" }
$csvInvlName = if ($csvInvalides)     { $csvInvalides.Name }     else { "" }
$csvChmName  = if ($csvChemins)       { $csvChemins.Name }       else { "" }
$csvDblName  = if ($csvDoublons)      { $csvDoublons.Name }      else { "" }
$csvExtName  = if ($csvExtensions)    { $csvExtensions.Name }    else { "" }
$csvPrmName  = if ($csvPermissions)   { $csvPermissions.Name }   else { "" }
$csvHOwName  = if ($csvHomeOwnership) { $csvHomeOwnership.Name } else { "" }
$csvHStName  = if ($csvHomeStats)     { $csvHomeStats.Name }     else { "" }

$cheminSource = ""
if ($dataInventaire.Count -gt 0) {
    $globalRow = $dataInventaire | Where-Object { $_.TypeLigne -eq "Global" } | Select-Object -First 1
    if ($globalRow) { $cheminSource = $globalRow.CheminAnalyse }
}
if (-not $cheminSource) { $cheminSource = $CheminOutput }

if ($fsMapping.Count -gt 0 -and -not $SplitByLevel1) {
    # ================================================================
    # MODE MAPPING — 1 rapport par FileShare + GLOBAL + INDEX
    # ================================================================
    Write-Log "Mode mapping — generation de $($fsMapping.Count) rapports par FileShare..." "INFO"

    $indexEntries = [System.Collections.Generic.List[hashtable]]::new()

    foreach ($fs in $fsMapping) {
        $nomFS     = $fs.NomFileShare
        $cheminUNC = $fs.CheminUNC
        if ([string]::IsNullOrWhiteSpace($nomFS)) { continue }

        Write-Log "  -> Rapport [$nomFS]..." "INFO"

        # Discover CSV files for this FS
        $csvSet = Get-FileShareCsvSet -OutputRoot $CheminOutput -NomFS $nomFS

        $fInv  = Load-FsData $csvSet.Inventaire    $dataInventaire    $nomFS $cheminUNC
        $fVol  = Load-FsData $csvSet.Volumineux     $dataVolumineux    $nomFS $cheminUNC
        $fChm  = Load-FsData $csvSet.Chemins        $dataChemins       $nomFS $cheminUNC
        $fExt  = Load-FsData $csvSet.Extensions     $dataExtensions    $nomFS $cheminUNC
        $fInvl = Load-FsData $csvSet.Invalides      $dataInvalides     $nomFS $cheminUNC
        $fDbl  = Load-FsData $csvSet.Doublons       $dataDoublons      $nomFS $cheminUNC
        $hasPermissionsCsv = ($null -ne $csvSet.Permissions) -or ($null -ne $csvPermissions)
        $fPrm  = if ($csvSet.Permissions) {
            @(Import-CsvSafe $csvSet.Permissions.FullName)
        } else {
            Get-PermissionsForScope -DataPermissions $dataPermissions -ScopeRoot $cheminUNC -NomFS $nomFS
        }
        $fHOw  = Load-FsData $csvSet.HomeOwnership  $dataHomeOwnership $nomFS $cheminUNC
        $fHSt  = Load-FsData $csvSet.HomeStats      $dataHomeStats     $nomFS $cheminUNC
        Write-Log "Section Permissions: $($fPrm.Count) lignes ACL trouvees pour $nomFS" "INFO"

        $fSid  = if ($csvSet.SidResolution)  { @(Import-CsvSafe $csvSet.SidResolution.FullName)  } else {
            @($dataSidResolution | Where-Object { $_.PSObject.Properties['NomFileShare'] -and $_.NomFileShare -eq $nomFS })
        }
        $fAD   = if ($csvSet.AccessDenied)   { @(Import-CsvSafe $csvSet.AccessDenied.FullName)   } else {
            @($dataAccessDenied  | Where-Object { $_.PSObject.Properties['NomFileShare'] -and $_.NomFileShare -eq $nomFS })
        }

        # Apply "Aucun" filter
        $fVol  = @($fVol  | Where-Object { $_.ResultatAnalyse -notmatch '^Aucun' })
        $fInvl = @($fInvl | Where-Object { $_.ResultatAnalyse -notmatch '^Aucun' -and $_.ProblemeType -ne "DoublePoint" })
        $fChm  = @($fChm  | Where-Object { $_.ResultatAnalyse -notmatch '^Aucun' })
        $fDbl  = @($fDbl  | Where-Object { $_.ResultatAnalyse -notmatch '^Aucun' })
        $fExt  = @($fExt  | Where-Object { $_.ResultatAnalyse -notmatch '^Aucun' })
        $fHOw  = @($fHOw  | Where-Object { $_.ResultatAnalyse -notmatch '^Aucun' })
        $fHSt  = @($fHSt  | Where-Object { $_.ResultatAnalyse -notmatch '^Aucun' })

        # Build stats for index
        $ligneTotal = $fInv | Where-Object { $_.DossierNiveau1 -eq "[TOTAL]" } | Select-Object -First 1
        $totFichiers = if ($ligneTotal -and $ligneTotal.NombreFichiers -match '^\d+$') { [long]$ligneTotal.NombreFichiers } else { 0 }
        $volGB = if ($ligneTotal) { [math]::Round((ConvertTo-LongSafe $ligneTotal.TailleOctets) / 1073741824.0, 2) } else { 0.0 }
        $blkErr = ($fInvl | Where-Object { $_.Severite -eq "ERROR" } | Measure-Object).Count + $fChm.Count
        # v3.3 — Extensions ne comptent plus dans les warnings
        $warns  = ($fInvl | Where-Object { $_.Severite -eq "WARN" } | Measure-Object).Count + $fVol.Count

        $subtitle  = "$($fs.TypeUsage) -> $($fs.CibleM365)"
        # v3.2 — Owner supprime de l'affichage (lecture conservee pour TypeUsage/CibleM365)
        $safeFsName = $nomFS -replace '[^\w-]', '_'
        $rptFile = Join-Path $reportOutputPath "Rapport_Assessment_PrimaGAZ_${safeFsName}_$timestamp.html"

        # v3.3 — Path du CSV detail pour pyramide des ages
        $fInvDetPath = if ($csvSet.InventaireDetail -and (Test-Path $csvSet.InventaireDetail.FullName)) {
            $csvSet.InventaireDetail.FullName
        } else { "" }

        try {
        $html = Build-FullReport `
            -Title         "Rapport Assessment - $nomFS" `
            -Subtitle      $subtitle `
            -Inventaire    $fInv `
            -Volumineux    $fVol `
            -Invalides     $fInvl `
            -Chemins       $fChm `
            -Doublons      $fDbl `
            -Extensions    $fExt `
            -Permissions   $fPrm `
            -SidResolution $fSid `
            -AccessDenied  $fAD `
            -HomeOwnership $fHOw `
            -HomeStats     $fHSt `
            -AgeSummaryPath $(if ($csvSet.AgeSummary) { $csvSet.AgeSummary.FullName } else { "" }) `
            -AgeDetails    $csvSet.AgeDetails `
            -InventaireDetailPath $fInvDetPath `
            -CheminUNCRoot $cheminUNC `
            -CsvVolName    $(if ($csvSet.Volumineux)  { $csvSet.Volumineux.Name }  else { $csvVolName  }) `
            -CsvInvName    $(if ($csvSet.Inventaire)  { $csvSet.Inventaire.Name }  else { $csvInvName  }) `
            -CsvInvalName  $(if ($csvSet.Invalides)   { $csvSet.Invalides.Name }   else { $csvInvlName }) `
            -CsvChmName    $(if ($csvSet.Chemins)     { $csvSet.Chemins.Name }     else { $csvChmName  }) `
            -CsvDblName    $(if ($csvSet.Doublons)    { $csvSet.Doublons.Name }    else { $csvDblName  }) `
            -CsvExtName    $(if ($csvSet.Extensions)  { $csvSet.Extensions.Name }  else { $csvExtName  }) `
            -CsvPrmName    $(if ($csvSet.Permissions) { $csvSet.Permissions.Name } else { $csvPrmName  }) `
            -CsvHOwName    $(if ($csvSet.HomeOwnership){ $csvSet.HomeOwnership.Name} else { $csvHOwName }) `
            -CsvHStName    $(if ($csvSet.HomeStats)   { $csvSet.HomeStats.Name }   else { $csvHStName  }) `
            -CheminSource  $cheminUNC `
            -NomFS         $nomFS `
            -OutputPath    $reportOutputPath `
            -HasPermissionsSource $hasPermissionsCsv

        [System.IO.File]::WriteAllText($rptFile, $html, [System.Text.UTF8Encoding]::new($false))
        Write-Log "[OK] Rapport $nomFS : $(Split-Path -Leaf $rptFile)" "SUCCESS"

        $indexEntries.Add(@{
            NomFS          = $nomFS
            TypeUsage      = $fs.TypeUsage
            CibleM365      = $fs.CibleM365
            Owner          = $fs.Owner
            EmailOwner     = $fs.EmailOwner
            VolumeGB       = $volGB
            BlockingErrors = $blkErr
            Warnings       = $warns
            TotalItems     = [math]::Max(1, $totFichiers)
            ReportFile     = (Split-Path -Leaf $rptFile)
        })
        } catch {
            Write-Log "[ERROR] Echec generation rapport $nomFS : $_" "ERROR"
        }
    }

    # ---- Rapport GLOBAL ----
    $globalFile = Join-Path $reportOutputPath "Rapport_Assessment_PrimaGAZ_GLOBAL_$timestamp.html"
    Write-Log "  -> Rapport global (confidentiel DSI)..." "INFO"
    try {
    $htmlGlobal = Build-FullReport `
        -Title         "Rapport Assessment Global PrimaGAZ (CONFIDENTIEL DSI)" `
        -Subtitle      "Migration FileShare vers M365 - Vue complete tous FileShares" `
        -Inventaire    $dataInventaire `
        -Volumineux    $dataVolumineux `
        -Invalides     $dataInvalides `
        -Chemins       $dataChemins `
        -Doublons      $dataDoublons `
        -Extensions    $dataExtensions `
        -Permissions   $dataPermissions `
        -SidResolution $dataSidResolution `
        -AccessDenied  $dataAccessDenied `
        -HomeOwnership $dataHomeOwnership `
        -HomeStats     $dataHomeStats `
        -AgeSummaryPath $(if (Test-Path (Join-Path $ageOutputPath "_SUMMARY.csv")) { Join-Path $ageOutputPath "_SUMMARY.csv" } else { "" }) `
        -CsvVolName    $csvVolName `
        -CsvInvName    $csvInvName `
        -CsvInvalName  $csvInvlName `
        -CsvChmName    $csvChmName `
        -CsvDblName    $csvDblName `
        -CsvExtName    $csvExtName `
        -CsvPrmName    $csvPrmName `
        -CsvHOwName    $csvHOwName `
        -CsvHStName    $csvHStName `
        -CheminSource  $cheminSource `
        -NomFS         "GLOBAL" `
        -OutputPath    $reportOutputPath `
        -HasPermissionsSource (($null -ne $csvPermissions) -or ($dataPermissions.Count -gt 0))
    [System.IO.File]::WriteAllText($globalFile, $htmlGlobal, [System.Text.UTF8Encoding]::new($false))
    Write-Log "[OK] Rapport global : $(Split-Path -Leaf $globalFile)" "SUCCESS"
    } catch {
        Write-Log "[ERROR] Echec generation rapport global : $_" "ERROR"
    }

    # ---- Index DSI ----
    $indexFile = Join-Path $reportOutputPath "Index_Assessment_PrimaGAZ_$timestamp.html"
    Write-Log "  -> Index DSI..." "INFO"
    try {
    $htmlIndex = Build-IndexReport `
        -Groups       $indexEntries `
        -GlobalFile   (Split-Path -Leaf $globalFile) `
        -FsMappings   $fsMapping
    [System.IO.File]::WriteAllText($indexFile, $htmlIndex, [System.Text.UTF8Encoding]::new($false))
    Write-Log "[OK] Index DSI : $(Split-Path -Leaf $indexFile)" "SUCCESS"
    } catch {
        Write-Log "[ERROR] Echec generation index DSI : $_" "ERROR"
    }

    Write-Log ""
    Write-Log "=== Fichiers generes dans $reportOutputPath ===" "SUCCESS"
    Write-Log "  Index DSI      : $(Split-Path -Leaf $indexFile)" "SUCCESS"
    Write-Log "  Rapport global : $(Split-Path -Leaf $globalFile)" "SUCCESS"
    Write-Log "  Rapports FS    : $($indexEntries.Count) fichier(s)" "SUCCESS"

} elseif ($SplitByLevel1) {
    # ================================================================
    # MODE SPLIT — Rapports par FileShare + N1 (mode legacy)
    # ================================================================
    Write-Log "Mode split active — generation des rapports par FileShare/N1..." "INFO"

    # Helper de filtrage inline
    filter Select-ByPrefix {
        param([string]$Col, [string]$Prefix)
        if ([string]::IsNullOrWhiteSpace($Prefix)) { return $_ }
        if ($_.$Col -like "$Prefix*") { return $_ }
    }

    # Identifier tous les chemins presents dans les CSV
    $allPathsSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($row in ($dataVolumineux + $dataInvalides + $dataChemins + $dataDoublons + $dataExtensions)) {
        if (-not [string]::IsNullOrWhiteSpace($row.CheminComplet)) { [void]$allPathsSet.Add($row.CheminComplet) }
    }
    foreach ($row in $dataPermissions) {
        if (-not [string]::IsNullOrWhiteSpace($row.CheminDossier)) { [void]$allPathsSet.Add($row.CheminDossier) }
    }
    foreach ($row in ($dataHomeOwnership + $dataHomeStats)) {
        if (-not [string]::IsNullOrWhiteSpace($row.CheminComplet)) { [void]$allPathsSet.Add($row.CheminComplet) }
    }

    $fsGroups    = [ordered]@{}
    $unmappedSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

    if ($fsMapping.Count -gt 0) {
        foreach ($path in $allPathsSet) {
            $fs = Get-FsForPath -Path $path -Mapping $fsMapping
            if ($fs) {
                $n1 = Get-N1Folder -Path $path -FsRoot $fs.CheminUNC
                if ([string]::IsNullOrWhiteSpace($n1)) { $n1 = "[RACINE]" }
                if ($fs.TypeUsage -eq "Personnel") { $n1 = "[TOUS-HOMES]" }
                $key = "$($fs.NomFileShare)|||$n1"
                if (-not $fsGroups.Contains($key)) {
                    $fsGroups[$key] = @{
                        NomFS     = $fs.NomFileShare
                        DossierN1 = $n1
                        TypeUsage = $fs.TypeUsage
                        Owner     = $fs.Owner
                        EmailOwner= $fs.EmailOwner
                        CibleM365 = $fs.CibleM365
                        FsRoot    = $fs.CheminUNC
                    }
                }
            } else {
                [void]$unmappedSet.Add($path)
            }
        }
    } else {
        # Sans mapping : detecter les N1 depuis l'inventaire
        foreach ($invRow in ($dataInventaire | Where-Object { $_.TypeLigne -eq "DossierNiveau1" })) {
            $n1  = $invRow.DossierNiveau1
            $key = "[FileShare]|||$n1"
            if (-not $fsGroups.Contains($key)) {
                $fsGroups[$key] = @{ NomFS="FileShare"; DossierN1=$n1; TypeUsage="Collaboratif"; Owner=""; EmailOwner=""; CibleM365="SharePoint"; FsRoot="" }
            }
        }
        # Completer avec les N1 detectes dans les chemins
        foreach ($path in $allPathsSet) {
            $parts = $path.TrimStart('\','/') -split '[/\\]', 3
            if ($parts.Count -ge 2 -and -not [string]::IsNullOrWhiteSpace($parts[1])) {
                $n1  = $parts[1]
                $key = "[FileShare]|||$n1"
                if (-not $fsGroups.Contains($key)) {
                    $fsGroups[$key] = @{ NomFS="FileShare"; DossierN1=$n1; TypeUsage="Collaboratif"; Owner=""; EmailOwner=""; CibleM365="SharePoint"; FsRoot="" }
                }
            }
        }
    }

    Write-Log "Groupes FS+N1 identifies : $($fsGroups.Count)" "INFO"

    $indexGroups = [System.Collections.Generic.List[hashtable]]::new()

    foreach ($keyGroup in $fsGroups.Keys) {
        $grp   = $fsGroups[$keyGroup]
        $nomFS = $grp.NomFS
        $n1    = $grp.DossierN1
        $fsRoot= $grp.FsRoot

        Write-Log "  -> Rapport [$nomFS / $n1]..." "INFO"

        $pfx = if ($fsRoot) {
            if ($n1 -in @("[RACINE]","[TOUS-HOMES]")) { $fsRoot.TrimEnd('\') }
            else { "$($fsRoot.TrimEnd('\'))\$n1" }
        } else { $null }

        $fVol  = if ($pfx) { @($dataVolumineux    | Where-Object { $_.CheminComplet  -like "$pfx*" }) } else { $dataVolumineux }
        $fInv  = if ($pfx) { @($dataInvalides     | Where-Object { $_.CheminComplet  -like "$pfx*" }) } else { $dataInvalides }
        $fChm  = if ($pfx) { @($dataChemins       | Where-Object { $_.CheminComplet  -like "$pfx*" }) } else { $dataChemins }
        $fDbl  = if ($pfx) { @($dataDoublons      | Where-Object { $_.CheminComplet  -like "$pfx*" }) } else { $dataDoublons }
        $fExt  = if ($pfx) { @($dataExtensions    | Where-Object { $_.CheminComplet  -like "$pfx*" }) } else { $dataExtensions }
        $hasPermissionsCsv = ($null -ne $csvPermissions)
        $fPrm  = if ($pfx) { Get-PermissionsForScope -DataPermissions $dataPermissions -ScopeRoot $pfx -NomFS $nomFS } else { @($dataPermissions) }
        $fHOw  = if ($pfx) { @($dataHomeOwnership | Where-Object { $_.CheminComplet  -like "$pfx*" }) } else { $dataHomeOwnership }
        $fHSt  = if ($pfx) { @($dataHomeStats     | Where-Object { $_.CheminComplet  -like "$pfx*" }) } else { $dataHomeStats }
        Write-Log "Section Permissions: $($fPrm.Count) lignes ACL trouvees pour $nomFS" "INFO"

        # Inventaire : ligne N1 specifique + ligne TOTAL recalculee
        $fInvN1 = @($dataInventaire | Where-Object { $_.DossierNiveau1 -eq $n1 -and $_.TypeLigne -eq "DossierNiveau1" })
        if ($fInvN1.Count -gt 0) {
            $totF  = ($fInvN1 | ForEach-Object { ConvertTo-LongSafe $_.NombreFichiers } | Measure-Object -Sum).Sum
            $totO  = ($fInvN1 | ForEach-Object { ConvertTo-LongSafe $_.TailleOctets } | Measure-Object -Sum).Sum
            $totTl = if ($totO -ge 1073741824) { "$([math]::Round($totO/1073741824.0,2)) GB" } elseif ($totO -ge 1048576) { "$([math]::Round($totO/1048576.0,1)) MB" } else { "$totO octets" }
            $totalRow = [PSCustomObject]@{ TypeLigne="Global"; DossierNiveau1="[TOTAL]"; CheminAnalyse=""; NombreFichiers=$totF; NombreDossiers=""; TailleOctets=$totO; TailleLisible=$totTl; DateAnalyse=""; ResultatAnalyse="Inventaire global" }
            $fInvN1 = @($totalRow) + $fInvN1
        }

        $blkErr = ($fInv | Where-Object { $_.Severite -eq "ERROR" } | Measure-Object).Count + $fChm.Count
        $warns  = $fExt.Count + ($fInv | Where-Object { $_.Severite -eq "WARN" } | Measure-Object).Count + $fVol.Count
        $volGB  = 0.0
        if ($fInvN1.Count -gt 0) {
            $tr = $fInvN1 | Where-Object { $_.DossierNiveau1 -eq "[TOTAL]" } | Select-Object -First 1
            if ($tr) { $volGB = (ConvertTo-LongSafe $tr.TailleOctets) / 1073741824.0 }
        }
        $totIt = [math]::Max(1, ($fInvN1 | Where-Object { $_.DossierNiveau1 -eq "[TOTAL]" } | Select-Object -ExpandProperty NombreFichiers -First 1) -as [long])

        $safeFsName = $nomFS -replace '[^\w-]', '_'
        $safeN1     = $n1    -replace '[^\w-]', '_'
        $rptFile    = Join-Path $reportOutputPath "Rapport_Assessment_${safeFsName}_${safeN1}_$timestamp.html"

        $reportSubtitle = "$(HtmlEnc $grp.TypeUsage) -> $(HtmlEnc $grp.CibleM365)"
        # v3.2 — Owner supprime de l'affichage

        $html = Build-FullReport `
            -Title         "Rapport Assessment $nomFS / $n1" `
            -Subtitle      $reportSubtitle `
            -Inventaire    $fInvN1 `
            -Volumineux    $fVol `
            -Invalides     $fInv `
            -Chemins       $fChm `
            -Doublons      $fDbl `
            -Extensions    $fExt `
            -Permissions   $fPrm `
            -SidResolution $dataSidResolution `
            -AccessDenied  $dataAccessDenied `
            -HomeOwnership $fHOw `
            -HomeStats     $fHSt `
            -CsvVolName    $csvVolName `
            -CsvInvalName  $csvInvlName `
            -CsvChmName    $csvChmName `
            -CsvDblName    $csvDblName `
            -CsvExtName    $csvExtName `
            -CsvPrmName    $csvPrmName `
            -CsvHOwName    $csvHOwName `
            -CsvHStName    $csvHStName `
            -CheminSource  $(if ($pfx) { $pfx } else { $CheminOutput }) `
            -CheminUNCRoot $(if ($pfx) { $pfx } else { "" }) `
            -NomFS         $nomFS `
            -OutputPath    $reportOutputPath `
            -HasPermissionsSource $hasPermissionsCsv

        [System.IO.File]::WriteAllText($rptFile, $html, [System.Text.UTF8Encoding]::new($false))
        Write-Log "  Rapport genere : $rptFile" "SUCCESS"

        $indexGroups.Add(@{
            NomFS         = $nomFS
            DossierN1     = $n1
            TypeUsage     = $grp.TypeUsage
            Owner         = $grp.Owner
            EmailOwner    = $grp.EmailOwner
            CibleM365     = $grp.CibleM365
            ReportFile    = (Split-Path -Leaf $rptFile)
            BlockingErrors= $blkErr
            Warnings      = $warns
            TotalItems    = $totIt
            VolumeGB      = $volGB
        })
    }

    # ---- Rapports complets par FileShare ----
    $fsByName = $indexGroups | Group-Object { $_.NomFS }
    foreach ($fs in $fsByName) {
        if ($fs.Group.Count -le 1) { continue }
        Write-Log "  -> Rapport complet FS [$($fs.Name)]..." "INFO"

        $fsMeta  = $fsMapping | Where-Object { $_.NomFileShare -eq $fs.Name } | Select-Object -First 1
        $fsRoot2 = if ($fsMeta -and $fsMeta.CheminUNC) { $fsMeta.CheminUNC.TrimEnd('\') } else { "" }

        $fVol2 = if ($fsRoot2) { @($dataVolumineux    | Where-Object { $_.CheminComplet  -like "$fsRoot2*" }) } else { $dataVolumineux }
        $fInv2 = if ($fsRoot2) { @($dataInvalides     | Where-Object { $_.CheminComplet  -like "$fsRoot2*" }) } else { $dataInvalides }
        $fChm2 = if ($fsRoot2) { @($dataChemins       | Where-Object { $_.CheminComplet  -like "$fsRoot2*" }) } else { $dataChemins }
        $fDbl2 = if ($fsRoot2) { @($dataDoublons      | Where-Object { $_.CheminComplet  -like "$fsRoot2*" }) } else { $dataDoublons }
        $fExt2 = if ($fsRoot2) { @($dataExtensions    | Where-Object { $_.CheminComplet  -like "$fsRoot2*" }) } else { $dataExtensions }
        $hasPermissionsCsv2 = ($null -ne $csvPermissions)
        $fPrm2 = if ($fsRoot2) { Get-PermissionsForScope -DataPermissions $dataPermissions -ScopeRoot $fsRoot2 -NomFS $fs.Name } else { @($dataPermissions) }
        $fHOw2 = if ($fsRoot2) { @($dataHomeOwnership | Where-Object { $_.CheminComplet  -like "$fsRoot2*" }) } else { $dataHomeOwnership }
        $fHSt2 = if ($fsRoot2) { @($dataHomeStats     | Where-Object { $_.CheminComplet  -like "$fsRoot2*" }) } else { $dataHomeStats }
        $fInvFS = @($dataInventaire | Where-Object { $fsRoot2 -eq "" -or $_.CheminAnalyse -like "$fsRoot2*" -or $_.DossierNiveau1 -eq "[TOTAL]" })
        Write-Log "Section Permissions: $($fPrm2.Count) lignes ACL trouvees pour $($fs.Name)" "INFO"

        $safeFsName2 = $fs.Name -replace '[^\w-]', '_'
        $rptFsFile   = Join-Path $reportOutputPath "Rapport_Assessment_${safeFsName2}_COMPLET_$timestamp.html"
        $cible2  = if ($fsMeta) { $fsMeta.CibleM365  } else { "" }
        # v3.2 — Owner supprime de l'affichage (variables owner2/email2 non necessaires)
        $usage2  = if ($fsMeta) { $fsMeta.TypeUsage   } else { "" }
        $sub2    = "$usage2 -> $cible2"

        $html2 = Build-FullReport `
            -Title         "Rapport Assessment $($fs.Name) (COMPLET)" `
            -Subtitle      "$sub2 | Confidentiel" `
            -Inventaire    $fInvFS `
            -Volumineux    $fVol2 `
            -Invalides     $fInv2 `
            -Chemins       $fChm2 `
            -Doublons      $fDbl2 `
            -Extensions    $fExt2 `
            -Permissions   $fPrm2 `
            -SidResolution $dataSidResolution `
            -AccessDenied  $dataAccessDenied `
            -HomeOwnership $fHOw2 `
            -HomeStats     $fHSt2 `
            -CsvVolName    $csvVolName `
            -CsvInvalName  $csvInvlName `
            -CsvChmName    $csvChmName `
            -CsvDblName    $csvDblName `
            -CsvExtName    $csvExtName `
            -CsvPrmName    $csvPrmName `
            -CsvHOwName    $csvHOwName `
            -CsvHStName    $csvHStName `
            -CheminSource  $fsRoot2 `
            -CheminUNCRoot $fsRoot2 `
            -NomFS         $($fs.Name) `
            -OutputPath    $reportOutputPath `
            -HasPermissionsSource $hasPermissionsCsv2

        [System.IO.File]::WriteAllText($rptFsFile, $html2, [System.Text.UTF8Encoding]::new($false))
        Write-Log "  Rapport FS complet : $rptFsFile" "SUCCESS"
    }

    # ---- Rapport GLOBAL ----
    $globalFile = Join-Path $reportOutputPath "Rapport_Assessment_PrimaGAZ_GLOBAL_$timestamp.html"
    Write-Log "  -> Rapport global (confidentiel DSI)..." "INFO"
    try {
    $htmlGlobal = Build-FullReport `
        -Title         "Rapport Assessment Global PrimaGAZ (CONFIDENTIEL DSI)" `
        -Subtitle      "Migration FileShare vers M365 - Vue complete tous FileShares" `
        -Inventaire    $dataInventaire `
        -Volumineux    $dataVolumineux `
        -Invalides     $dataInvalides `
        -Chemins       $dataChemins `
        -Doublons      $dataDoublons `
        -Extensions    $dataExtensions `
        -Permissions   $dataPermissions `
        -SidResolution $dataSidResolution `
        -AccessDenied  $dataAccessDenied `
        -HomeOwnership $dataHomeOwnership `
        -HomeStats     $dataHomeStats `
        -AgeSummaryPath $(if (Test-Path (Join-Path $ageOutputPath "_SUMMARY.csv")) { Join-Path $ageOutputPath "_SUMMARY.csv" } else { "" }) `
        -CsvVolName    $csvVolName `
        -CsvInvName    $csvInvName `
        -CsvInvalName  $csvInvlName `
        -CsvChmName    $csvChmName `
        -CsvDblName    $csvDblName `
        -CsvExtName    $csvExtName `
        -CsvPrmName    $csvPrmName `
        -CsvHOwName    $csvHOwName `
        -CsvHStName    $csvHStName `
        -CheminSource  $cheminSource `
        -NomFS         "GLOBAL" `
        -OutputPath    $reportOutputPath `
        -HasPermissionsSource (($null -ne $csvPermissions) -or ($dataPermissions.Count -gt 0))
    [System.IO.File]::WriteAllText($globalFile, $htmlGlobal, [System.Text.UTF8Encoding]::new($false))
    Write-Log "[OK] Rapport global : $(Split-Path -Leaf $globalFile)" "SUCCESS"
    } catch {
        Write-Log "[ERROR] Echec generation rapport global : $_" "ERROR"
    }

    # ---- Index DSI ----
    $indexFile = Join-Path $reportOutputPath "Index_Assessment_PrimaGAZ_$timestamp.html"
    Write-Log "  -> Index DSI comparatif..." "INFO"
    try {
    $htmlIndex = Build-IndexReport `
        -Groups        $indexGroups `
        -GlobalFile    (Split-Path -Leaf $globalFile) `
        -FsMappings    $fsMapping `
        -UnmappedPaths @($unmappedSet)
    [System.IO.File]::WriteAllText($indexFile, $htmlIndex, [System.Text.UTF8Encoding]::new($false))
    Write-Log "[OK] Index DSI : $(Split-Path -Leaf $indexFile)" "SUCCESS"
    } catch {
        Write-Log "[ERROR] Echec generation index DSI : $_" "ERROR"
    }

    Write-Log ""
    Write-Log "=== Fichiers generes dans $reportOutputPath ===" "SUCCESS"
    Write-Log "  Index DSI      : $(Split-Path -Leaf $indexFile)" "SUCCESS"
    Write-Log "  Rapport global : $(Split-Path -Leaf $globalFile)" "SUCCESS"
    Write-Log "  Rapports N1    : $($indexGroups.Count) fichier(s)" "SUCCESS"

} else {
    # ================================================================
    # MODE STANDARD — Rapport global unique (retrocompatible)
    # ================================================================
    $rapportFile = Join-Path $reportOutputPath "Rapport_Assessment_PrimaGAZ_$timestamp.html"
    Write-Log "Mode standard — generation du rapport global..." "INFO"

    try {
    $html = Build-FullReport `
        -Title         "Rapport Assessment PrimaGAZ" `
        -Subtitle      "Migration FileShare vers M365" `
        -Inventaire    $dataInventaire `
        -Volumineux    $dataVolumineux `
        -Invalides     $dataInvalides `
        -Chemins       $dataChemins `
        -Doublons      $dataDoublons `
        -Extensions    $dataExtensions `
        -Permissions   $dataPermissions `
        -SidResolution $dataSidResolution `
        -AccessDenied  $dataAccessDenied `
        -HomeOwnership $dataHomeOwnership `
        -HomeStats     $dataHomeStats `
        -AgeSummaryPath $(if (Test-Path (Join-Path $ageOutputPath "_SUMMARY.csv")) { Join-Path $ageOutputPath "_SUMMARY.csv" } else { "" }) `
        -CsvVolName    $csvVolName `
        -CsvInvName    $csvInvName `
        -CsvInvalName  $csvInvlName `
        -CsvChmName    $csvChmName `
        -CsvDblName    $csvDblName `
        -CsvExtName    $csvExtName `
        -CsvPrmName    $csvPrmName `
        -CsvHOwName    $csvHOwName `
        -CsvHStName    $csvHStName `
        -CheminSource  $cheminSource `
        -NomFS         "PrimaGAZ" `
        -OutputPath    $reportOutputPath `
        -HasPermissionsSource (($null -ne $csvPermissions) -or ($dataPermissions.Count -gt 0))

    [System.IO.File]::WriteAllText($rapportFile, $html, [System.Text.UTF8Encoding]::new($false))
    Write-Log "[OK] Rapport genere : $rapportFile" "SUCCESS"
    Write-Log "Ouvrez le fichier dans un navigateur pour visualiser le rapport." "INFO"
    } catch {
        Write-Log "[ERROR] Echec generation rapport global : $_" "ERROR"
    }
}

Write-Log "Generation terminee." "SUCCESS"
