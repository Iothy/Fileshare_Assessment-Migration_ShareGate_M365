<#
.SYNOPSIS
    Module de gestion centralisée des sorties pour les scripts d'assessment PrimaGAZ.
.DESCRIPTION
    Centralise la logique de création et de navigation dans la structure :
    Output/<Scope>/<yyyyMMdd_HHmmss>/ puis un sous-dossier par source.
#>

function ConvertTo-PlainHashtable {
    param($InputObject)

    if ($null -eq $InputObject) { return $null }
    if ($InputObject -is [hashtable]) { return $InputObject }
    if ($InputObject -is [System.Collections.IDictionary]) {
        $table = @{}
        foreach ($key in $InputObject.Keys) {
            $table[$key] = ConvertTo-PlainHashtable -InputObject $InputObject[$key]
        }
        return $table
    }
    if ($InputObject -is [pscustomobject]) {
        $table = @{}
        foreach ($property in $InputObject.PSObject.Properties) {
            $table[$property.Name] = ConvertTo-PlainHashtable -InputObject $property.Value
        }
        return $table
    }
    if ($InputObject -is [System.Collections.IEnumerable] -and $InputObject -isnot [string]) {
        return @(
            foreach ($item in $InputObject) {
                ConvertTo-PlainHashtable -InputObject $item
            }
        )
    }
    return $InputObject
}

function Ensure-Directory {
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path -Path $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
    return $Path
}

function New-AssessmentRun {
    param(
        [Parameter(Mandatory)] [string]$Scope,
        [Parameter(Mandatory)] [string]$BaseOutput,
        [string]$FileSharePath = ''
    )

    $ts = Get-Date -Format 'yyyyMMdd_HHmmss'
    $scopePath = Ensure-Directory -Path (Join-Path $BaseOutput $Scope)
    $runPath = Ensure-Directory -Path (Join-Path $scopePath $ts)

    return [PSCustomObject]@{
        Id            = "${Scope}_${ts}"
        Path          = $runPath
        Csv           = $runPath
        Metadata      = $runPath
        Logs          = $runPath
        Errors        = $runPath
        Manifest      = Join-Path $runPath 'manifest.json'
        RootLog       = Join-Path $runPath 'execution.log'
        Timestamp     = $ts
        Scope         = $Scope
        FileSharePath = $FileSharePath
        StartedAt     = Get-Date
    }
}

function _Reconstruct-Run {
    param(
        [string]$RunPath,
        [string]$Scope
    )

    $ts = Split-Path $RunPath -Leaf
    $fileSharePath = ''
    $manifestPath = Join-Path $RunPath 'manifest.json'

    if (Test-Path $manifestPath) {
        try {
            $content = [System.IO.File]::ReadAllText($manifestPath, [System.Text.Encoding]::UTF8)
            $obj = $content | ConvertFrom-Json
            if ($obj.scope -and $obj.scope.fileshare_path) {
                $fileSharePath = $obj.scope.fileshare_path
            }
        }
        catch {
            Write-Warning "_Reconstruct-Run : échec de lecture du manifest '$manifestPath' — $_"
        }
    }

    return [PSCustomObject]@{
        Id            = "${Scope}_${ts}"
        Path          = $RunPath
        Csv           = $RunPath
        Metadata      = $RunPath
        Logs          = $RunPath
        Errors        = $RunPath
        Manifest      = $manifestPath
        RootLog       = Join-Path $RunPath 'execution.log'
        Timestamp     = $ts
        Scope         = $Scope
        FileSharePath = $fileSharePath
        StartedAt     = Get-Date
    }
}

function Get-LatestRun {
    param(
        [Parameter(Mandatory)] [string]$Scope,
        [Parameter(Mandatory)] [string]$BaseOutput
    )

    $scopePath = Join-Path $BaseOutput $Scope
    if (-not (Test-Path $scopePath)) { return $null }

    $latest = Get-ChildItem -Path $scopePath -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -match '^\d{8}_\d{6}$' } |
        Sort-Object Name -Descending |
        Select-Object -First 1

    if (-not $latest) { return $null }
    return _Reconstruct-Run -RunPath $latest.FullName -Scope $Scope
}

function Get-AllRuns {
    param(
        [Parameter(Mandatory)] [string]$Scope,
        [Parameter(Mandatory)] [string]$BaseOutput
    )

    $scopePath = Join-Path $BaseOutput $Scope
    if (-not (Test-Path $scopePath)) { return @() }

    $dirs = Get-ChildItem -Path $scopePath -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -match '^\d{8}_\d{6}$' } |
        Sort-Object Name -Descending

    return @(
        foreach ($dir in $dirs) {
            _Reconstruct-Run -RunPath $dir.FullName -Scope $Scope
        }
    )
}

function Get-AssessmentSourceFolder {
    param(
        [Parameter(Mandatory)] $Run,
        [Parameter(Mandatory)] [string]$SourceIdentifier
    )

    return Ensure-Directory -Path (Join-Path $Run.Path $SourceIdentifier)
}

function Get-AssessmentSourceFilePath {
    param(
        [Parameter(Mandatory)] $Run,
        [Parameter(Mandatory)] [string]$SourceIdentifier,
        [Parameter(Mandatory)] [string]$Prefix,
        [Parameter(Mandatory)] [string]$Extension
    )

    $folder = Get-AssessmentSourceFolder -Run $Run -SourceIdentifier $SourceIdentifier
    return Join-Path $folder ("{0}_{1}.{2}" -f $Prefix, $SourceIdentifier, $Extension.TrimStart('.'))
}

function Get-AssessmentRootFilePath {
    param(
        [Parameter(Mandatory)] $Run,
        [Parameter(Mandatory)] [string]$FileName
    )

    return Join-Path $Run.Path $FileName
}

function Write-EmptyCsv {
    param(
        [Parameter(Mandatory)] [string]$Path,
        [Parameter(Mandatory)] [array]$Columns
    )

    $header = ($Columns | ForEach-Object { '"{0}"' -f (($_ -as [string]) -replace '"', '""') }) -join ';'
    [System.IO.File]::WriteAllText($Path, $header + [Environment]::NewLine, [System.Text.UTF8Encoding]::new($true))
}

function Write-RunManifest {
    param(
        [Parameter(Mandatory)] $Run,
        [string]$Status = 'Success',
        [array]$Scripts = @(),
        [hashtable]$Summary = @{}
    )

    $existing = @{}
    if (Test-Path $Run.Manifest) {
        try {
            $existing = ConvertTo-PlainHashtable -InputObject ((Get-Content -Path $Run.Manifest -Raw -Encoding UTF8) | ConvertFrom-Json)
        }
        catch {
            $existing = @{}
        }
    }

    $now = Get-Date
    $durationMin = [math]::Round(($now - $Run.StartedAt).TotalMinutes, 1)
    $currentUser = try {
        if ($PSVersionTable.PSEdition -eq 'Desktop') {
            [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
        }
        else {
            whoami 2>$null
        }
    }
    catch {
        if ($env:USERNAME) { "$env:USERDOMAIN\$env:USERNAME" } else { 'Inconnu' }
    }
    if (-not $currentUser) { $currentUser = if ($env:USERNAME) { $env:USERNAME } else { 'Inconnu' } }

    $mergedSummary = @{}
    if ($existing.ContainsKey('summary') -and $existing.summary) {
        $mergedSummary = ConvertTo-PlainHashtable -InputObject $existing.summary
    }
    foreach ($entry in $Summary.GetEnumerator()) {
        $mergedSummary[$entry.Key] = $entry.Value
    }

    $manifest = [ordered]@{
        run_id    = "$($Run.Scope)_$($Run.Timestamp)"
        scope     = [ordered]@{
            name           = $Run.Scope
            fileshare_path = $Run.FileSharePath
        }
        execution = [ordered]@{
            started_at         = $Run.StartedAt.ToString('o')
            ended_at           = $now.ToString('o')
            duration_minutes   = $durationMin
            host               = $env:COMPUTERNAME
            user               = $currentUser
            powershell_version = "$($PSVersionTable.PSVersion)"
            scripts_version    = 'v4.0'
        }
        status    = [ordered]@{
            global = $Status
        }
        scripts   = @($Scripts | Where-Object { $null -ne $_ })
        summary   = $mergedSummary
    }

    foreach ($property in @('mapping', 'sources', 'reports', 'files', 'warnings', 'errors', 'controls', 'completeness')) {
        if ($existing.ContainsKey($property) -and -not $manifest.Contains($property)) {
            $manifest[$property] = $existing[$property]
        }
    }

    try {
        $json = $manifest | ConvertTo-Json -Depth 8
        [System.IO.File]::WriteAllText($Run.Manifest, $json, [System.Text.Encoding]::UTF8)
    }
    catch {
        Write-Warning "Write-RunManifest : impossible d'écrire $($Run.Manifest) — $_"
    }
}

function Update-RunsHistory {
    param(
        [Parameter(Mandatory)] $Run,
        [string]$Status = 'Success'
    )

    $scopePath = Split-Path $Run.Path -Parent
    $runsFile = Join-Path $scopePath '_runs.json'
    $now = Get-Date
    $durationMin = [math]::Round(($now - $Run.StartedAt).TotalMinutes, 1)

    $entry = [ordered]@{
        id               = $Run.Id
        timestamp        = $Run.Timestamp
        status           = $Status
        duration_minutes = $durationMin
        path             = $Run.Path
    }

    $runs = New-Object 'System.Collections.Generic.List[object]'
    if (Test-Path $runsFile) {
        try {
            $content = [System.IO.File]::ReadAllText($runsFile, [System.Text.Encoding]::UTF8)
            $existing = $content | ConvertFrom-Json
            if ($existing) {
                foreach ($r in $existing) { $runs.Add($r) | Out-Null }
            }
        }
        catch {
            Write-Warning "Update-RunsHistory : impossible de lire $runsFile — $_"
        }
    }

    $runs.Add($entry) | Out-Null

    try {
        $json = $runs.ToArray() | ConvertTo-Json -Depth 3
        [System.IO.File]::WriteAllText($runsFile, $json, [System.Text.Encoding]::UTF8)
    }
    catch {
        Write-Warning "Update-RunsHistory : impossible d'écrire $runsFile — $_"
    }
}

function Initialize-ScriptOutput {
    param(
        [Parameter(Mandatory)] $Run,
        [Parameter(Mandatory)] [string]$ScriptName
    )

    return [PSCustomObject]@{
        CsvPath      = Join-Path $Run.Path "$ScriptName.csv"
        MetadataPath = Join-Path $Run.Path "$ScriptName.json"
        LogPath      = Join-Path $Run.Path "$ScriptName.log"
        ErrorsPath   = $Run.Path
    }
}

Export-ModuleMember -Function New-AssessmentRun, Get-LatestRun, Get-AllRuns,
    Write-RunManifest, Update-RunsHistory, Initialize-ScriptOutput,
    Get-AssessmentSourceFolder, Get-AssessmentSourceFilePath, Get-AssessmentRootFilePath,
    Write-EmptyCsv
