<#
.SYNOPSIS
    Module de gestion centralisée des sorties pour les scripts d'assessment PrimaGAZ.
.DESCRIPTION
    Centralise toute la logique de création et de navigation dans la structure de dossiers
    hiérarchique : Output/<Scope>/<yyyyMMdd_HHmmss>/{csv,metadata,logs,errors}/
.NOTES
    Projet  : PrimaGAZ - Migration FileShare vers M365
    Phase   : 01 - Assessment
    Version : v1.0
#>

function New-AssessmentRun {
    <#
    .SYNOPSIS
        Crée une nouvelle structure de run et retourne l'objet Run.
    #>
    param(
        [Parameter(Mandatory)] [string]$Scope,
        [Parameter(Mandatory)] [string]$BaseOutput,
        [string]$FileSharePath = ""
    )

    $ts = Get-Date -Format "yyyyMMdd_HHmmss"
    $runPath = Join-Path $BaseOutput (Join-Path $Scope $ts)

    $subFolders = @("csv", "metadata", "logs", "errors")
    foreach ($sub in $subFolders) {
        $subPath = Join-Path $runPath $sub
        if (-not (Test-Path $subPath)) {
            try { New-Item -ItemType Directory -Path $subPath -Force | Out-Null }
            catch { Write-Warning "New-AssessmentRun : échec de création du dossier '$subPath' — $_" }
        }
    }

    $run = [PSCustomObject]@{
        Id            = "${Scope}_${ts}"
        Path          = $runPath
        Csv           = Join-Path $runPath "csv"
        Metadata      = Join-Path $runPath "metadata"
        Logs          = Join-Path $runPath "logs"
        Errors        = Join-Path $runPath "errors"
        Manifest      = Join-Path $runPath "manifest.json"
        Timestamp     = $ts
        Scope         = $Scope
        FileSharePath = $FileSharePath
        StartedAt     = (Get-Date)
    }

    return $run
}

function _Reconstruct-Run {
    <#
    .SYNOPSIS
        Reconstruit un objet Run à partir d'un chemin de dossier existant.
    #>
    param(
        [string]$RunPath,
        [string]$Scope
    )
    $ts = Split-Path $RunPath -Leaf
    $fileSharePath = ""
    $manifestPath = Join-Path $RunPath "manifest.json"

    if (Test-Path $manifestPath) {
        try {
            $content = [System.IO.File]::ReadAllText($manifestPath, [System.Text.Encoding]::UTF8)
            $obj = $content | ConvertFrom-Json
            if ($obj.scope -and $obj.scope.fileshare_path) {
                $fileSharePath = $obj.scope.fileshare_path
            }
        } catch { Write-Warning "_Reconstruct-Run : échec de lecture du manifest '$manifestPath' — $_" }
    }

    return [PSCustomObject]@{
        Id            = "${Scope}_${ts}"
        Path          = $RunPath
        Csv           = Join-Path $RunPath "csv"
        Metadata      = Join-Path $RunPath "metadata"
        Logs          = Join-Path $RunPath "logs"
        Errors        = Join-Path $RunPath "errors"
        Manifest      = $manifestPath
        Timestamp     = $ts
        Scope         = $Scope
        FileSharePath = $fileSharePath
        StartedAt     = (Get-Date)
    }
}

function Get-LatestRun {
    <#
    .SYNOPSIS
        Retourne le dernier run pour un scope donné, ou $null si aucun.
    #>
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
    <#
    .SYNOPSIS
        Retourne tous les runs pour un scope, du plus récent au plus ancien.
    #>
    param(
        [Parameter(Mandatory)] [string]$Scope,
        [Parameter(Mandatory)] [string]$BaseOutput
    )

    $scopePath = Join-Path $BaseOutput $Scope
    if (-not (Test-Path $scopePath)) { return @() }

    $dirs = Get-ChildItem -Path $scopePath -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -match '^\d{8}_\d{6}$' } |
        Sort-Object Name -Descending

    if (-not $dirs) { return @() }

    $runs = [System.Collections.Generic.List[object]]::new()
    foreach ($dir in $dirs) {
        $runs.Add((_Reconstruct-Run -RunPath $dir.FullName -Scope $Scope))
    }
    return $runs.ToArray()
}

function Write-RunManifest {
    <#
    .SYNOPSIS
        Écrit le manifest.json dans le dossier de run.
    #>
    param(
        [Parameter(Mandatory)] $Run,
        [string]$Status = "Success",
        [array]$Scripts = @(),
        [hashtable]$Summary = @{}
    )

    $now = (Get-Date)
    $durationMin = [math]::Round(($now - $Run.StartedAt).TotalMinutes, 1)

    $currentUser = try {
        if ($PSVersionTable.PSEdition -eq 'Desktop') {
            [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
        } else {
            (whoami 2>$null)
        }
    } catch {
        if ($env:USERNAME) { "$env:USERDOMAIN\$env:USERNAME" } else { "Inconnu" }
    }
    if (-not $currentUser) {
        $currentUser = if ($env:USERNAME) { $env:USERNAME } else { "Inconnu" }
    }

    $scriptsArr = @()
    foreach ($s in $Scripts) {
        if ($null -ne $s) { $scriptsArr += $s }
    }

    $manifest = [ordered]@{
        run_id    = "$($Run.Scope)_$($Run.Timestamp)"
        scope     = [ordered]@{
            name           = $Run.Scope
            fileshare_path = $Run.FileSharePath
        }
        execution = [ordered]@{
            started_at         = $Run.StartedAt.ToString("o")
            ended_at           = $now.ToString("o")
            duration_minutes   = $durationMin
            host               = $env:COMPUTERNAME
            user               = $currentUser
            powershell_version = "$($PSVersionTable.PSVersion)"
            scripts_version    = "v3.0"
        }
        status    = [ordered]@{
            global = $Status
        }
        scripts = $scriptsArr
        summary = $Summary
    }

    try {
        $json = $manifest | ConvertTo-Json -Depth 5
        $utf8 = [System.Text.Encoding]::UTF8
        [System.IO.File]::WriteAllText($Run.Manifest, $json, $utf8)
    } catch {
        Write-Warning "Write-RunManifest : impossible d'écrire $($Run.Manifest) — $_"
    }
}

function Update-RunsHistory {
    <#
    .SYNOPSIS
        Met à jour le fichier _runs.json du scope avec l'entrée du run courant.
    #>
    param(
        [Parameter(Mandatory)] $Run,
        [string]$Status = "Success"
    )

    $scopePath  = Split-Path $Run.Path -Parent
    $runsFile   = Join-Path $scopePath "_runs.json"

    $now = (Get-Date)
    $durationMin = [math]::Round(($now - $Run.StartedAt).TotalMinutes, 1)

    $entry = [ordered]@{
        id               = $Run.Id
        timestamp        = $Run.Timestamp
        status           = $Status
        duration_minutes = $durationMin
        path             = $Run.Path
    }

    $runs = [System.Collections.Generic.List[object]]::new()

    if (Test-Path $runsFile) {
        try {
            $content = [System.IO.File]::ReadAllText($runsFile, [System.Text.Encoding]::UTF8)
            $existing = $content | ConvertFrom-Json
            if ($existing) {
                foreach ($r in $existing) { $runs.Add($r) }
            }
        } catch {
            Write-Warning "Update-RunsHistory : impossible de lire $runsFile — $_"
        }
    }

    $runs.Add($entry)

    try {
        $json = $runs.ToArray() | ConvertTo-Json -Depth 3
        $utf8 = [System.Text.Encoding]::UTF8
        [System.IO.File]::WriteAllText($runsFile, $json, $utf8)
    } catch {
        Write-Warning "Update-RunsHistory : impossible d'écrire $runsFile — $_"
    }
}

function Initialize-ScriptOutput {
    <#
    .SYNOPSIS
        Retourne les chemins de sortie pour un script donné dans le contexte d'un Run.
    #>
    param(
        [Parameter(Mandatory)] $Run,
        [Parameter(Mandatory)] [string]$ScriptName
    )

    return [PSCustomObject]@{
        CsvPath      = Join-Path $Run.Csv "$ScriptName.csv"
        MetadataPath = Join-Path $Run.Metadata "$ScriptName.json"
        LogPath      = Join-Path $Run.Logs "$ScriptName.log"
        ErrorsPath   = $Run.Errors
    }
}

Export-ModuleMember -Function New-AssessmentRun, Get-LatestRun, Get-AllRuns,
                                Write-RunManifest, Update-RunsHistory, Initialize-ScriptOutput
