function Test-FileShareAssessmentPrerequisite {
    <#
    .SYNOPSIS
    Vérifie que l'hôte et le périmètre sont prêts avant un assessment.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$ConfigurationPath,

        [Parameter()]
        [ValidateRange(1, 10240)]
        [int]$MinimumFreeSpaceGB = 5
    )

    $configuration = Import-FileShareAssessmentConfiguration -ConfigurationPath $ConfigurationPath
    $checks = [System.Collections.Generic.List[object]]::new()
    $windowsHost = $env:OS -eq 'Windows_NT'
    $checks.Add([PSCustomObject]@{
            Name = 'Windows'
            Status = if ($windowsHost) { 'Pass' } else { 'Fail' }
            Message = if ($windowsHost) { 'Hôte Windows détecté.' } else { 'Un assessment SMB doit être exécuté sur Windows.' }
        })
    $checks.Add([PSCustomObject]@{
            Name = 'PowerShell'
            Status = if ($PSVersionTable.PSVersion.Major -ge 5) { 'Pass' } else { 'Fail' }
            Message = "PowerShell $($PSVersionTable.PSVersion) détecté."
        })

    $outputPath = $configuration.Assessment.OutputRoot
    $outputParent = $outputPath
    while (-not (Test-Path -LiteralPath $outputParent) -and (Split-Path -Path $outputParent -Parent) -ne $outputParent) {
        $outputParent = Split-Path -Path $outputParent -Parent
    }
    try {
        $drive = Get-Item -LiteralPath $outputParent -ErrorAction Stop | Select-Object -ExpandProperty PSDrive
        $freeSpaceGB = [math]::Round($drive.Free / 1GB, 2)
        $checks.Add([PSCustomObject]@{
                Name = 'OutputFreeSpace'
                Status = if ($freeSpaceGB -ge $MinimumFreeSpaceGB) { 'Pass' } else { 'Fail' }
                Message = "$freeSpaceGB GB libres sur '$($drive.Root)' (minimum : $MinimumFreeSpaceGB GB)."
            })
    }
    catch {
        $checks.Add([PSCustomObject]@{ Name = 'OutputFreeSpace'; Status = 'Fail'; Message = "Impossible de vérifier l'espace de sortie : $($_.Exception.Message)" })
    }

    try {
        $probeFile = Join-Path -Path $outputParent -ChildPath ('.assessment-write-probe-{0}.tmp' -f [guid]::NewGuid())
        [System.IO.File]::WriteAllText($probeFile, '')
        Remove-Item -LiteralPath $probeFile -Force -ErrorAction Stop
        $checks.Add([PSCustomObject]@{ Name = 'OutputWrite'; Status = 'Pass'; Message = "Le répertoire de sortie '$outputParent' est accessible en écriture." })
    }
    catch {
        $checks.Add([PSCustomObject]@{ Name = 'OutputWrite'; Status = 'Fail'; Message = "Le répertoire de sortie '$outputParent' n'est pas accessible en écriture : $($_.Exception.Message)" })
    }

    if ($windowsHost -and (Test-Path -LiteralPath $outputParent)) {
        try {
            $broadWrite = @(Get-Acl -LiteralPath $outputParent -ErrorAction Stop).Access | Where-Object {
                $_.AccessControlType -eq 'Allow' -and
                $_.IdentityReference -match 'Everyone|Builtin\\Users|Authenticated Users' -and
                $_.FileSystemRights.ToString() -match 'Write|Modify|FullControl'
            }
            $checks.Add([PSCustomObject]@{
                    Name = 'OutputSecurity'
                    Status = if ($broadWrite) { 'Warn' } else { 'Pass' }
                    Message = if ($broadWrite) { "Les ACL de '$outputParent' autorisent une écriture étendue ; utilisez un emplacement restreint." } else { "Aucune autorisation d'écriture étendue détectée sur '$outputParent'." }
                })
        }
        catch {
            $checks.Add([PSCustomObject]@{ Name = 'OutputSecurity'; Status = 'Warn'; Message = "Impossible d'inspecter les ACL de sortie : $($_.Exception.Message)" })
        }
    }

    $mapping = @(Import-Csv -LiteralPath $configuration.Assessment.MappingCsv -Delimiter ';' -Encoding UTF8)
    foreach ($entry in $mapping) {
        if ([string]::IsNullOrWhiteSpace($entry.CheminUNC)) { continue }
        try {
            $item = Get-Item -LiteralPath $entry.CheminUNC.Trim() -ErrorAction Stop
            $checks.Add([PSCustomObject]@{ Name = 'SourceAccess'; Target = $entry.CheminUNC; Status = if ($item.PSIsContainer) { 'Pass' } else { 'Fail' }; Message = if ($item.PSIsContainer) { 'Chemin source accessible en lecture.' } else { "Le chemin source n'est pas un dossier." } })
        }
        catch {
            $checks.Add([PSCustomObject]@{ Name = 'SourceAccess'; Target = $entry.CheminUNC; Status = 'Fail'; Message = "Chemin source inaccessible : $($_.Exception.Message)" })
        }
    }

    [PSCustomObject]@{
        IsReady = @($checks | Where-Object Status -eq 'Fail').Count -eq 0
        Checks = $checks.ToArray()
    }
}
