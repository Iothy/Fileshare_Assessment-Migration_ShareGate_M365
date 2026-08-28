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
    if ([string]::IsNullOrWhiteSpace($outputParent)) {
        $checks.Add([PSCustomObject]@{ Name = 'OutputPath'; Status = 'Fail'; Message = "Le répertoire de sortie '$outputPath' ne peut pas être résolu." })
        return [PSCustomObject]@{
            IsReady = $false
            Checks = $checks.ToArray()
        }
    }
    if ($outputParent -match '^\\\\') {
        $checks.Add([PSCustomObject]@{ Name = 'OutputFreeSpace'; Status = 'Warn'; Message = "L'espace libre du partage UNC '$outputParent' ne peut pas être déterminé de manière fiable ; vérifiez au moins $MinimumFreeSpaceGB GB libres avec son administrateur." })
    }
    else {
        try {
            if ($windowsHost) {
                $driveRoot = (Split-Path -Path $outputParent -Qualifier) + '\'
                $freeSpaceGB = [math]::Round(([System.IO.DriveInfo]::new($driveRoot).AvailableFreeSpace / 1GB), 2)
            }
            else {
                $drive = Get-Item -LiteralPath $outputParent -ErrorAction Stop | Select-Object -ExpandProperty PSDrive
                $driveRoot = $drive.Root
                $freeSpaceGB = [math]::Round($drive.Free / 1GB, 2)
            }
            $checks.Add([PSCustomObject]@{
                    Name = 'OutputFreeSpace'
                    Status = if ($freeSpaceGB -ge $MinimumFreeSpaceGB) { 'Pass' } else { 'Fail' }
                    Message = "$freeSpaceGB GB libres sur '$driveRoot' (minimum : $MinimumFreeSpaceGB GB)."
                })
        }
        catch {
            $checks.Add([PSCustomObject]@{ Name = 'OutputFreeSpace'; Status = 'Fail'; Message = "Impossible de vérifier l'espace de sortie : $($_.Exception.Message)" })
        }
    }

    if (-not (Test-Path -LiteralPath $outputParent)) {
        $checks.Add([PSCustomObject]@{ Name = 'OutputWrite'; Status = 'Fail'; Message = "Aucun répertoire parent existant n'a été trouvé pour la sortie '$outputPath'." })
    }
    else {
        $probeFile = Join-Path -Path $outputParent -ChildPath ('.assessment-write-probe-{0}.tmp' -f [guid]::NewGuid())
        try {
            [System.IO.File]::WriteAllText($probeFile, '')
            $checks.Add([PSCustomObject]@{ Name = 'OutputWrite'; Status = 'Pass'; Message = "Le répertoire de sortie '$outputParent' est accessible en écriture." })
        }
        catch {
            $checks.Add([PSCustomObject]@{ Name = 'OutputWrite'; Status = 'Fail'; Message = "Le répertoire de sortie '$outputParent' n'est pas accessible en écriture : $($_.Exception.Message)" })
        }
        finally {
            if (Test-Path -LiteralPath $probeFile) {
                try {
                    Remove-Item -LiteralPath $probeFile -Force -ErrorAction Stop
                }
                catch {
                    $checks.Add([PSCustomObject]@{ Name = 'OutputProbeCleanup'; Status = 'Warn'; Message = "Le fichier de test '$probeFile' n'a pas pu être supprimé : $($_.Exception.Message)" })
                }
            }
        }
    }

    if ($windowsHost -and (Test-Path -LiteralPath $outputParent)) {
        try {
            $broadWrite = (Get-Acl -LiteralPath $outputParent -ErrorAction Stop).Access | Where-Object {
                try {
                    $identitySid = $_.IdentityReference.Translate([System.Security.Principal.SecurityIdentifier]).Value
                }
                catch {
                    $identitySid = ''
                }
                $rights = [uint32]$_.FileSystemRights
                $hasWriteRights = (($rights -band [uint32][System.Security.AccessControl.FileSystemRights]::Write) -ne 0) -or
                    (($rights -band [uint32][System.Security.AccessControl.FileSystemRights]::Modify) -eq [uint32][System.Security.AccessControl.FileSystemRights]::Modify) -or
                    (($rights -band [uint32][System.Security.AccessControl.FileSystemRights]::FullControl) -eq [uint32][System.Security.AccessControl.FileSystemRights]::FullControl)
                $_.AccessControlType -eq 'Allow' -and
                $identitySid -in @('S-1-1-0', 'S-1-5-11', 'S-1-5-32-545') -and
                $hasWriteRights
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
