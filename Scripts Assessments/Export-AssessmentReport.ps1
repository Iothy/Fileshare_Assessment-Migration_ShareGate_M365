<#
.SYNOPSIS
    Génère les rapports HTML d'assessment par source et les synthèses globales.
.DESCRIPTION
    Produit un rapport HTML par source dans son sous-dossier, un rapport global et un index
    à la racine du run. Le script consomme le mapping simplifié 6 colonnes via le module
    FileShareAssessment et n'affiche que des synthèses agrégées.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$CheminOutput = (Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Definition) 'Output'),

    [Parameter(Mandatory = $false, HelpMessage = 'Chemin du fichier FileShareMapping.csv')]
    [string]$FileShareMapping = '',

    [Parameter(Mandatory = $false, HelpMessage = 'Dossier de destination des rapports HTML')]
    [string]$ReportOutputPath = '',

    [Parameter(Mandatory = $false, HelpMessage = 'Paramètre conservé pour compatibilité, désormais ignoré.')]
    [switch]$SplitByLevel1,

    [Parameter(Mandatory = $false)]
    [PSCustomObject]$Run = $null
)

Import-Module (Join-Path -Path $PSScriptRoot -ChildPath 'Modules/PrimaGAZ.Assessment.psm1') -Force
$outputModulePath = Join-Path -Path $PSScriptRoot -ChildPath 'Modules/PrimaGAZ.Output.psm1'
if (Test-Path $outputModulePath) { Import-Module $outputModulePath -Force }
$fileShareAssessmentModulePath = Join-Path -Path $PSScriptRoot -ChildPath 'FileShareAssessment/FileShareAssessment.psd1'
if (Test-Path $fileShareAssessmentModulePath) { Import-Module $fileShareAssessmentModulePath -Force }

if ($Run) {
    $CheminOutput = $Run.Path
    $ReportOutputPath = $Run.Path
}
elseif (-not [string]::IsNullOrWhiteSpace($ReportOutputPath)) {
    $CheminOutput = $ReportOutputPath
}

if (-not (Test-Path -Path $CheminOutput)) {
    throw "Le dossier '$CheminOutput' n'existe pas."
}

if ($SplitByLevel1) {
    Write-Log 'Le paramètre -SplitByLevel1 est ignoré : les rapports sont désormais générés uniquement par source.' 'WARN'
}

function HtmlEnc {
    param([AllowNull()]$Value)
    return [System.Net.WebUtility]::HtmlEncode([string]$Value)
}

function Import-CsvSafe {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -Path $Path)) { return @() }
    try {
        return @(Import-Csv -Path $Path -Delimiter ';' -Encoding UTF8)
    }
    catch {
        Write-Log "Lecture CSV impossible : $Path - $($_.Exception.Message)" 'WARN'
        return @()
    }
}

function Get-RelativePathFromRoot {
    param(
        [Parameter(Mandatory)][string]$RootPath,
        [Parameter(Mandatory)][string]$Path
    )

    $rootFull = [System.IO.Path]::GetFullPath($RootPath).TrimEnd([System.IO.Path]::DirectorySeparatorChar, [System.IO.Path]::AltDirectorySeparatorChar)
    $pathFull = [System.IO.Path]::GetFullPath($Path)
    if ($pathFull.StartsWith($rootFull, [System.StringComparison]::OrdinalIgnoreCase)) {
        return $pathFull.Substring($rootFull.Length).TrimStart([System.IO.Path]::DirectorySeparatorChar, [System.IO.Path]::AltDirectorySeparatorChar)
    }
    return Split-Path -Leaf $pathFull
}

function Get-CsvRowCount {
    param([string]$Path)
    if (-not (Test-Path -Path $Path)) { return 0 }
    try { return @(Import-CsvSafe -Path $Path).Count } catch { return 0 }
}

function Get-ExecutionLogStatus {
    param([string]$Path)
    if (-not (Test-Path -Path $Path)) { return 'Missing' }
    $content = Get-Content -Path $Path -Encoding UTF8 -ErrorAction SilentlyContinue
    $statusLine = $content | Where-Object { $_ -match '^Status\s*:' } | Select-Object -Last 1
    if (-not $statusLine) { return 'Unknown' }
    return ($statusLine -split ':', 2)[1].Trim()
}

function Convert-ToInt64 {
    param($Value)
    $parsed = 0L
    if ([long]::TryParse([string]$Value, [ref]$parsed)) { return $parsed }
    return 0L
}

function Format-ByteValue {
    param([long]$Value)
    if ($Value -ge 1TB) { return '{0:N2} TB' -f ($Value / 1TB) }
    if ($Value -ge 1GB) { return '{0:N2} GB' -f ($Value / 1GB) }
    if ($Value -ge 1MB) { return '{0:N2} MB' -f ($Value / 1MB) }
    if ($Value -ge 1KB) { return '{0:N2} KB' -f ($Value / 1KB) }
    return '{0:N0} octets' -f $Value
}

function Get-SourceRowSet {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Identifier
    )

    $rows = Import-CsvSafe -Path $Path
    foreach ($row in $rows) {
        if ($row.PSObject.Properties.Name -notcontains 'SourceIdentifier') {
            Add-Member -InputObject $row -MemberType NoteProperty -Name SourceIdentifier -Value $Identifier -Force
        }
    }
    return @($rows)
}

function Get-InventorySummary {
    param([array]$InventoryRows, [array]$SummaryRows)

    if ($SummaryRows.Count -gt 0) {
        $row = $SummaryRows | Select-Object -First 1
        return [PSCustomObject]@{
            Files       = Convert-ToInt64 $row.NombreFichiers
            Folders     = Convert-ToInt64 $row.NombreDossiers
            SizeBytes   = Convert-ToInt64 $row.TailleOctets
            SizeLabel   = if ($row.TailleLisible) { [string]$row.TailleLisible } else { Format-ByteValue -Value (Convert-ToInt64 $row.TailleOctets) }
            DetailFiles = Convert-ToInt64 $row.FichiersDetail
        }
    }

    $totalRow = $InventoryRows | Where-Object { $_.TypeLigne -eq 'Global' -or $_.DossierNiveau1 -eq '[TOTAL]' } | Select-Object -First 1
    if ($totalRow) {
        $sizeBytes = Convert-ToInt64 $totalRow.TailleOctets
        return [PSCustomObject]@{
            Files       = Convert-ToInt64 $totalRow.NombreFichiers
            Folders     = Convert-ToInt64 $totalRow.NombreDossiers
            SizeBytes   = $sizeBytes
            SizeLabel   = if ($totalRow.TailleLisible) { [string]$totalRow.TailleLisible } else { Format-ByteValue -Value $sizeBytes }
            DetailFiles = 0
        }
    }

    return [PSCustomObject]@{ Files = 0; Folders = 0; SizeBytes = 0; SizeLabel = '0 octet'; DetailFiles = 0 }
}

function Build-StatusBadge {
    param([string]$Status)
    $css = switch -Regex ($Status) {
        '^(?i:success|ok)$' { 'success' }
        '^(?i:partial)' { 'warning' }
        '^(?i:missing|failed)' { 'danger' }
        '^(?i:skipped)' { 'warning' }
        default { 'neutral' }
    }
    return '<span class="badge {0}">{1}</span>' -f $css, (HtmlEnc $Status)
}

function Get-PathTooLongStatus {
    param(
        [string]$Path,
        [array]$Rows
    )

    if (-not (Test-Path $Path)) { return 'Failed' }
    if (@($Rows).Count -eq 0) { return 'Success' }
    if (@($Rows | Where-Object { $_.StatutControle -eq 'Skipped' }).Count -eq @($Rows).Count) { return 'Skipped' }
    return 'Partial'
}

function Build-HtmlDocument {
    param(
        [Parameter(Mandatory)][string]$Title,
        [Parameter(Mandatory)][string]$BodyHtml
    )

    @"
<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="utf-8" />
<title>$(HtmlEnc $Title)</title>
<style>
body{font-family:Segoe UI,Arial,sans-serif;margin:24px;color:#1f2937;background:#f8fafc}
h1,h2,h3{color:#111827}table{border-collapse:collapse;width:100%;margin:16px 0;background:#fff}
th,td{border:1px solid #d1d5db;padding:8px;text-align:left;vertical-align:top}th{background:#e5e7eb}.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:12px}.card{background:#fff;border:1px solid #d1d5db;border-radius:8px;padding:12px}.muted{color:#6b7280}.badge{display:inline-block;padding:2px 8px;border-radius:999px;font-size:12px;font-weight:600}.badge.success{background:#dcfce7;color:#166534}.badge.warning{background:#fef3c7;color:#92400e}.badge.danger{background:#fee2e2;color:#991b1b}.badge.neutral{background:#e5e7eb;color:#374151}code{background:#eef2ff;padding:2px 4px;border-radius:4px}a{color:#2563eb}ul{padding-left:18px}
</style>
</head>
<body>
$BodyHtml
</body>
</html>
"@
}

function Build-PerSourceReport {
    param(
        [Parameter(Mandatory)]$Source,
        [Parameter(Mandatory)][hashtable]$FileMap,
        [Parameter(Mandatory)]$Summary,
        [Parameter(Mandatory)]$ControlStatuses,
        [Parameter(Mandatory)][string]$RunId,
        [Parameter(Mandatory)][datetime]$RunDate,
        [Parameter(Mandatory)][string]$ScriptVersion,
        [Parameter(Mandatory)][string]$CompletenessStatus,
        [Parameter(Mandatory)][int]$WarningCount,
        [Parameter(Mandatory)][int]$ErrorCount
    )

    $targetDisplay = if ([string]::IsNullOrWhiteSpace($Source.TargetFolder)) { $Source.TargetSPOURL } else { '{0}/{1}' -f $Source.TargetSPOURL, $Source.TargetFolder }
    $fileRows = foreach ($key in $FileMap.Keys | Sort-Object) {
        $path = $FileMap[$key]
        $fileName = Split-Path -Leaf $path
        $count = if ($fileName -like '*.csv' -and (Test-Path $path)) { Get-CsvRowCount -Path $path } else { 0 }
        "<tr><td>$(HtmlEnc $key)</td><td><a href=`"$(HtmlEnc $fileName)`">$(HtmlEnc $fileName)</a></td><td>$(HtmlEnc $count)</td></tr>"
    }

    $statusRows = foreach ($controlName in $ControlStatuses.Keys | Sort-Object) {
        "<tr><td>$(HtmlEnc $controlName)</td><td>$(Build-StatusBadge -Status $ControlStatuses[$controlName])</td></tr>"
    }

    $body = @"
<h1>Rapport $(HtmlEnc $Source.SourceIdentifier)</h1>
<p class="muted">Source : <code>$(HtmlEnc $Source.SourcePath)</code></p>
<div class="grid">
  <div class="card"><strong>RunId</strong><br/>$(HtmlEnc $RunId)</div>
  <div class="card"><strong>Date du run</strong><br/>$(HtmlEnc ($RunDate.ToString('yyyy-MM-dd HH:mm:ss')))</div>
  <div class="card"><strong>Version script</strong><br/>$(HtmlEnc $ScriptVersion)</div>
  <div class="card"><strong>Statut</strong><br/>$(Build-StatusBadge -Status $CompletenessStatus)</div>
</div>
<h2>Mapping cible</h2>
<table>
<tr><th>Champ</th><th>Valeur</th></tr>
<tr><td>TargetType</td><td>$(HtmlEnc $Source.TargetType)</td></tr>
<tr><td>TargetSPOURL</td><td>$(HtmlEnc $Source.TargetSPOURL)</td></tr>
<tr><td>TargetFolder</td><td>$(HtmlEnc $Source.TargetFolder)</td></tr>
<tr><td>DateFilter</td><td>$(HtmlEnc $Source.DateFilter)</td></tr>
<tr><td>Permissions</td><td>$(HtmlEnc $Source.Permissions)</td></tr>
<tr><td>Destination normalisée</td><td>$(HtmlEnc $targetDisplay)</td></tr>
</table>
<h2>Métadonnées dérivées</h2>
<table>
<tr><th>Champ</th><th>Valeur</th></tr>
<tr><td>ServerFqdn</td><td>$(HtmlEnc $Source.ServerFqdn)</td></tr>
<tr><td>ServerShort</td><td>$(HtmlEnc $Source.ServerShort)</td></tr>
<tr><td>ShareName</td><td>$(HtmlEnc $Source.ShareName)</td></tr>
<tr><td>RelativePath</td><td>$(HtmlEnc $Source.RelativePath)</td></tr>
<tr><td>LeafName</td><td>$(HtmlEnc $Source.LeafName)</td></tr>
<tr><td>SourceIdentifier</td><td>$(HtmlEnc $Source.SourceIdentifier)</td></tr>
</table>
<h2>Synthèse</h2>
<div class="grid">
  <div class="card"><strong>Fichiers</strong><br/>$(HtmlEnc $Summary.Files)</div>
  <div class="card"><strong>Dossiers</strong><br/>$(HtmlEnc $Summary.Folders)</div>
  <div class="card"><strong>Volume</strong><br/>$(HtmlEnc $Summary.SizeLabel)</div>
  <div class="card"><strong>Warnings</strong><br/>$(HtmlEnc $WarningCount)</div>
  <div class="card"><strong>Errors</strong><br/>$(HtmlEnc $ErrorCount)</div>
</div>
<h2>Statut des contrôles</h2>
<table>
<tr><th>Contrôle</th><th>Statut</th></tr>
$($statusRows -join "`n")
</table>
<h2>Fichiers disponibles</h2>
<p class="muted">Les CSV complets sont fournis en téléchargement ; le HTML ne reprend volontairement que les agrégats.</p>
<table>
<tr><th>Type</th><th>Fichier</th><th>Lignes</th></tr>
$($fileRows -join "`n")
</table>
"@

    return Build-HtmlDocument -Title "Rapport $($Source.SourceIdentifier)" -BodyHtml $body
}

function Build-IndexReport {
    param(
        [Parameter(Mandatory)][array]$Entries,
        [Parameter(Mandatory)][string]$RunId,
        [Parameter(Mandatory)][datetime]$RunDate,
        [Parameter(Mandatory)][string]$GlobalReportFile
    )

    $rows = foreach ($entry in $Entries) {
        $target = if ([string]::IsNullOrWhiteSpace($entry.Source.TargetFolder)) { $entry.Source.TargetSPOURL } else { '{0}/{1}' -f $entry.Source.TargetSPOURL, $entry.Source.TargetFolder }
        "<tr><td><code>$(HtmlEnc $entry.Source.SourcePath)</code></td><td>$(HtmlEnc $target)</td><td>$(Build-StatusBadge -Status $entry.Status)</td><td>$(HtmlEnc $entry.Summary.SizeLabel)</td><td>$(HtmlEnc $entry.WarningCount)</td><td>$(HtmlEnc $entry.ErrorCount)</td><td><a href=`"./$(HtmlEnc $entry.Source.SourceIdentifier)/Rapport_$(HtmlEnc $entry.Source.SourceIdentifier).html`">Ouvrir</a></td></tr>"
    }

    $body = @"
<h1>Index_Assessment</h1>
<p class="muted">RunId : $(HtmlEnc $RunId) — $(HtmlEnc ($RunDate.ToString('yyyy-MM-dd HH:mm:ss')))</p>
<p><a href="$(HtmlEnc $GlobalReportFile)">Ouvrir le rapport global</a></p>
<table>
<tr><th>Source</th><th>Cible</th><th>Statut</th><th>Volumétrie</th><th>Warnings</th><th>Errors</th><th>Rapport</th></tr>
$($rows -join "`n")
</table>
"@

    return Build-HtmlDocument -Title 'Index Assessment' -BodyHtml $body
}

function Build-GlobalReport {
    param(
        [Parameter(Mandatory)][array]$Entries,
        [Parameter(Mandatory)][string]$RunId,
        [Parameter(Mandatory)][datetime]$RunDate,
        [Parameter(Mandatory)][string]$ScriptVersion
    )

    $totalFiles = ($Entries | ForEach-Object { $_.Summary.Files } | Measure-Object -Sum).Sum
    $totalFolders = ($Entries | ForEach-Object { $_.Summary.Folders } | Measure-Object -Sum).Sum
    $totalBytes = ($Entries | ForEach-Object { $_.Summary.SizeBytes } | Measure-Object -Sum).Sum
    $warningCount = ($Entries | ForEach-Object { $_.WarningCount } | Measure-Object -Sum).Sum
    $errorCount = ($Entries | ForEach-Object { $_.ErrorCount } | Measure-Object -Sum).Sum

    $rows = foreach ($entry in $Entries) {
        "<tr><td>$(HtmlEnc $entry.Source.SourceIdentifier)</td><td><code>$(HtmlEnc $entry.Source.SourcePath)</code></td><td>$(Build-StatusBadge -Status $entry.Status)</td><td>$(HtmlEnc $entry.Summary.Files)</td><td>$(HtmlEnc $entry.Summary.Folders)</td><td>$(HtmlEnc $entry.Summary.SizeLabel)</td></tr>"
    }

    $body = @"
<h1>Rapport_Global</h1>
<p class="muted">RunId : $(HtmlEnc $RunId) — $(HtmlEnc ($RunDate.ToString('yyyy-MM-dd HH:mm:ss')))</p>
<div class="grid">
  <div class="card"><strong>Sources</strong><br/>$(HtmlEnc $Entries.Count)</div>
  <div class="card"><strong>Fichiers</strong><br/>$(HtmlEnc $totalFiles)</div>
  <div class="card"><strong>Dossiers</strong><br/>$(HtmlEnc $totalFolders)</div>
  <div class="card"><strong>Volume total</strong><br/>$(HtmlEnc (Format-ByteValue -Value $totalBytes))</div>
  <div class="card"><strong>Warnings</strong><br/>$(HtmlEnc $warningCount)</div>
  <div class="card"><strong>Errors</strong><br/>$(HtmlEnc $errorCount)</div>
</div>
<p class="muted">Version script : $(HtmlEnc $ScriptVersion)</p>
<table>
<tr><th>Identifiant</th><th>Source</th><th>Statut</th><th>Fichiers</th><th>Dossiers</th><th>Volume</th></tr>
$($rows -join "`n")
</table>
"@

    return Build-HtmlDocument -Title 'Rapport Global' -BodyHtml $body
}

$runRoot = if ($Run) { $Run.Path } else { $CheminOutput }
$runDate = Get-Date
$runId = if ($Run) { $Run.Id } else { 'Assessment_' + (Get-Date -Format 'yyyyMMdd_HHmmss') }
$scriptVersion = 'v4.0'

$mappingValidation = $null
$mappingRows = @()
if (-not [string]::IsNullOrWhiteSpace($FileShareMapping) -and (Test-Path -Path $FileShareMapping)) {
    $mappingValidation = Test-FileShareMapping -Path $FileShareMapping
    foreach ($warning in $mappingValidation.Warnings) {
        Write-Log ('Mapping ligne {0}: {1}' -f $warning.LineNumber, $warning.Message) 'WARN'
    }
    if ($mappingValidation.IsValid) {
        $mappingRows = @($mappingValidation.Rows)
    }
}

if ($mappingRows.Count -eq 0) {
    $mappingRows = @(
        foreach ($dir in Get-ChildItem -Path $runRoot -Directory -ErrorAction SilentlyContinue) {
            [PSCustomObject]@{
                SourcePath       = $dir.Name
                TargetType       = ''
                TargetSPOURL     = ''
                TargetFolder     = ''
                DateFilter       = ''
                Permissions      = ''
                ServerFqdn       = ''
                ServerShort      = ''
                ShareName        = ''
                RelativePath     = ''
                LeafName         = $dir.Name
                SourceIdentifier = $dir.Name
                LineNumber       = 0
            }
        }
    )
}

$entries = New-Object 'System.Collections.Generic.List[object]'

foreach ($source in $mappingRows) {
    $sourceFolder = Join-Path $runRoot $source.SourceIdentifier
    if (-not (Test-Path -Path $sourceFolder)) {
        $null = New-Item -ItemType Directory -Path $sourceFolder -Force
    }

    $fileMap = [ordered]@{
        Synthese        = Join-Path $sourceFolder ("Synthese_{0}.csv" -f $source.SourceIdentifier)
        Inventaire      = Join-Path $sourceFolder ("Inventaire_{0}.csv" -f $source.SourceIdentifier)
        InventaireDetail= Join-Path $sourceFolder ("InventaireDetail_{0}.csv" -f $source.SourceIdentifier)
        Permissions     = Join-Path $sourceFolder ("Permissions_{0}.csv" -f $source.SourceIdentifier)
        CheminsLongs    = Join-Path $sourceFolder ("CheminsLongs_{0}.csv" -f $source.SourceIdentifier)
        Extensions      = Join-Path $sourceFolder ("Extensions_{0}.csv" -f $source.SourceIdentifier)
        Doublons        = Join-Path $sourceFolder ("Doublons_{0}.csv" -f $source.SourceIdentifier)
        AccesRefuses    = Join-Path $sourceFolder ("AccesRefuses_{0}.csv" -f $source.SourceIdentifier)
    }

    $summaryRows = Get-SourceRowSet -Path $fileMap.Synthese -Identifier $source.SourceIdentifier
    $inventoryRows = Get-SourceRowSet -Path $fileMap.Inventaire -Identifier $source.SourceIdentifier
    $cheminsRows = Get-SourceRowSet -Path $fileMap.CheminsLongs -Identifier $source.SourceIdentifier
    $extensionsRows = Get-SourceRowSet -Path $fileMap.Extensions -Identifier $source.SourceIdentifier
    $doublonsRows = Get-SourceRowSet -Path $fileMap.Doublons -Identifier $source.SourceIdentifier
    $accessRows = Get-SourceRowSet -Path $fileMap.AccesRefuses -Identifier $source.SourceIdentifier

    $summary = Get-InventorySummary -InventoryRows $inventoryRows -SummaryRows $summaryRows
    $warningCount = @($extensionsRows).Count + @($doublonsRows).Count + @($accessRows | Where-Object { $_.TypeErreur -eq 'AccessDenied' }).Count
    $pathTooLongIssueCount = @($cheminsRows | Where-Object { $_.StatutControle -ne 'Skipped' }).Count
    $errorCount = $pathTooLongIssueCount + @($accessRows | Where-Object { $_.TypeErreur -and $_.TypeErreur -ne 'AccessDenied' }).Count

    $controlStatuses = [ordered]@{
        Inventaire       = if (Test-Path $fileMap.Inventaire) { 'Success' } else { 'Failed' }
        Permissions      = if (Test-Path $fileMap.Permissions) { 'Success' } else { 'Failed' }
        Extensions       = if (Test-Path $fileMap.Extensions) { if ((Get-CsvRowCount $fileMap.Extensions) -gt 0) { 'Warning' } else { 'Success' } } else { 'Failed' }
        Doublons         = if (Test-Path $fileMap.Doublons) { if ((Get-CsvRowCount $fileMap.Doublons) -gt 0) { 'Warning' } else { 'Success' } } else { 'Failed' }
        CheminsLongs     = Get-PathTooLongStatus -Path $fileMap.CheminsLongs -Rows $cheminsRows
        AccesRefuses     = if (Test-Path $fileMap.AccesRefuses) { if ((Get-CsvRowCount $fileMap.AccesRefuses) -gt 0) { 'Warning' } else { 'Success' } } else { 'Failed' }
    }

    $statuses = @($controlStatuses.Values)
    if ($statuses -contains 'Failed') {
        $overallStatus = 'Partial'
    }
    elseif ($statuses -contains 'Partial' -or $statuses -contains 'Warning') {
        $overallStatus = 'Partial'
    }
    else {
        $overallStatus = 'Success'
    }

    $reportPath = Join-Path $sourceFolder ("Rapport_{0}.html" -f $source.SourceIdentifier)
    $fileMap.Rapport = $reportPath

    $reportHtml = Build-PerSourceReport -Source $source -FileMap $fileMap -Summary $summary -ControlStatuses $controlStatuses -RunId $runId -RunDate $runDate -ScriptVersion $scriptVersion -CompletenessStatus $overallStatus -WarningCount $warningCount -ErrorCount $errorCount
    [System.IO.File]::WriteAllText($reportPath, $reportHtml, [System.Text.UTF8Encoding]::new($false))

    $executionSummaryPath = Join-Path $sourceFolder ("Execution_{0}.log" -f $source.SourceIdentifier)
    @"
RunId              : $runId
SourcePath         : $($source.SourcePath)
SourceIdentifier   : $($source.SourceIdentifier)
TargetType         : $($source.TargetType)
TargetSPOURL       : $($source.TargetSPOURL)
TargetFolder       : $($source.TargetFolder)
Warnings           : $warningCount
Errors             : $errorCount
Status             : $overallStatus
"@ | Out-File -FilePath $executionSummaryPath -Encoding UTF8
    $fileMap.Execution = $executionSummaryPath

    $producedFiles = @(
        foreach ($file in $fileMap.Values) {
            if (Test-Path -Path $file) {
                [PSCustomObject]@{
                    path       = (Get-RelativePathFromRoot -RootPath $runRoot -Path $file)
                    row_count  = if ($file -like '*.csv') { Get-CsvRowCount -Path $file } else { $null }
                }
            }
        }
    )

    $entries.Add([PSCustomObject]@{
        Source        = $source
        Folder        = $sourceFolder
        Summary       = $summary
        Status        = $overallStatus
        WarningCount  = $warningCount
        ErrorCount    = $errorCount
        ControlStatus = $controlStatuses
        ProducedFiles = $producedFiles
        FileMap       = $fileMap
    }) | Out-Null
}

$globalReportPath = Join-Path $runRoot 'Rapport_Global.html'
$indexReportPath = Join-Path $runRoot 'Index_Assessment.html'

$globalHtml = Build-GlobalReport -Entries $entries.ToArray() -RunId $runId -RunDate $runDate -ScriptVersion $scriptVersion
[System.IO.File]::WriteAllText($globalReportPath, $globalHtml, [System.Text.UTF8Encoding]::new($false))

$indexHtml = Build-IndexReport -Entries $entries.ToArray() -RunId $runId -RunDate $runDate -GlobalReportFile './Rapport_Global.html'
[System.IO.File]::WriteAllText($indexReportPath, $indexHtml, [System.Text.UTF8Encoding]::new($false))

$manifestStatus = if (($entries | Where-Object Status -eq 'Partial').Count -gt 0) { 'Partial' } else { 'Success' }
if (($entries | Where-Object { $_.Status -eq 'Success' }).Count -eq 0 -and $entries.Count -gt 0) {
    $manifestStatus = 'Failed'
}

$manifest = [ordered]@{
    run_id = $runId
    mapping = [ordered]@{
        file = $FileShareMapping
        is_valid = if ($mappingValidation) { [bool]$mappingValidation.IsValid } else { $null }
        warnings = if ($mappingValidation) { @($mappingValidation.Warnings) } else { @() }
        errors = if ($mappingValidation) { @($mappingValidation.Errors) } else { @() }
    }
    execution = [ordered]@{
        generated_at = $runDate.ToString('o')
        script_version = $scriptVersion
    }
    status = [ordered]@{
        global = $manifestStatus
    }
    reports = [ordered]@{
        index = 'Index_Assessment.html'
        global = 'Rapport_Global.html'
    }
    sources = @(
        foreach ($entry in $entries) {
            [ordered]@{
                source_path = $entry.Source.SourcePath
                source_identifier = $entry.Source.SourceIdentifier
                mapping = [ordered]@{
                    SourcePath = $entry.Source.SourcePath
                    TargetType = $entry.Source.TargetType
                    TargetSPOURL = $entry.Source.TargetSPOURL
                    TargetFolder = $entry.Source.TargetFolder
                    DateFilter = $entry.Source.DateFilter
                    Permissions = $entry.Source.Permissions
                }
                derived = [ordered]@{
                    ServerFqdn = $entry.Source.ServerFqdn
                    ServerShort = $entry.Source.ServerShort
                    ShareName = $entry.Source.ShareName
                    RelativePath = $entry.Source.RelativePath
                    LeafName = $entry.Source.LeafName
                    SourceIdentifier = $entry.Source.SourceIdentifier
                }
                summary = [ordered]@{
                    files = $entry.Summary.Files
                    folders = $entry.Summary.Folders
                    size_bytes = $entry.Summary.SizeBytes
                    size_label = $entry.Summary.SizeLabel
                    warnings = $entry.WarningCount
                    errors = $entry.ErrorCount
                    access_denied = (Get-CsvRowCount -Path $entry.FileMap.AccesRefuses)
                }
                controls = $entry.ControlStatus
                status = $entry.Status
                files = $entry.ProducedFiles
            }
        }
    )
}

[System.IO.File]::WriteAllText((Join-Path $runRoot 'manifest.json'), ($manifest | ConvertTo-Json -Depth 8), [System.Text.Encoding]::UTF8)

@"
RunId              : $runId
GeneratedAt        : $($runDate.ToString('yyyy-MM-dd HH:mm:ss'))
Sources            : $($entries.Count)
GlobalStatus       : $manifestStatus
IndexReport        : Index_Assessment.html
GlobalReport       : Rapport_Global.html
"@ | Out-File -FilePath (Join-Path $runRoot 'execution.log') -Encoding UTF8

if ($Run) {
    Write-RunManifest -Run $Run -Status $manifestStatus -Summary @{
        sources = $entries.Count
        reports = 2
    }
}

return $globalReportPath
