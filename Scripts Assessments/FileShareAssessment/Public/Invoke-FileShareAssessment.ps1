function Invoke-FileShareAssessment {
    <#
    .SYNOPSIS
    Lance un assessment complet à partir d'une configuration explicite.

    .DESCRIPTION
    Crée un run horodaté, exécute les scripts historiques activés dans l'ordre défini,
    puis écrit un manifest de suivi sans réécrire les algorithmes métier existants.

    .PARAMETER ConfigurationPath
    Chemin vers le fichier JSON de configuration.

    .PARAMETER Credential
    Credential SMB optionnel à transmettre aux scripts historiques qui l'acceptent.

    .PARAMETER Server
    Nom du serveur SMB à transmettre aux scripts historiques qui l'acceptent.

    .PARAMETER PassThru
    Retourne le détail du run et des contrôles exécutés.

    .EXAMPLE
    Invoke-FileShareAssessment -ConfigurationPath '.\Config\FileShareAssessment.json'
    #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$ConfigurationPath,

        [Parameter()]
        [System.Management.Automation.PSCredential]
        [System.Management.Automation.Credential()]
        $Credential,

        [Parameter()]
        [string]$Server,

        [Parameter()]
        [switch]$PassThru
    )

    $configuration = Import-FileShareAssessmentConfiguration -ConfigurationPath $ConfigurationPath
    $enabledControls = @($configuration.Controls | Where-Object Enabled)
    $effectiveServer = if ($PSBoundParameters.ContainsKey('Server')) { $Server } else { $configuration.Assessment.Server }
    $plan = [PSCustomObject]@{
        Name       = $configuration.Assessment.Name
        Scope      = $configuration.Assessment.Scope
        MappingCsv = $configuration.Assessment.MappingCsv
        OutputRoot = $configuration.Assessment.OutputRoot
        Controls   = @($enabledControls.Name)
    }

    if (-not $PSCmdlet.ShouldProcess($configuration.Assessment.Scope, "Exécuter $($enabledControls.Count) contrôle(s) d'assessment")) {
        return $plan
    }

    $prerequisites = Test-FileShareAssessmentPrerequisite -ConfigurationPath $ConfigurationPath
    if (-not $prerequisites.IsReady) {
        $failures = @($prerequisites.Checks | Where-Object Status -eq 'Fail' | ForEach-Object { $_.Message }) -join [Environment]::NewLine
        throw [System.InvalidOperationException]::new("Le préflight a échoué. Corrigez les éléments suivants avant de lancer l'assessment :$([Environment]::NewLine)$failures")
    }
    foreach ($warning in @($prerequisites.Checks | Where-Object Status -eq 'Warn')) {
        Write-Warning $warning.Message
    }

    $outputModulePath = Join-Path -Path $script:AssessmentRoot -ChildPath 'Modules/PrimaGAZ.Output.psm1'
    Import-Module -Name $outputModulePath -Force -ErrorAction Stop

    $run = New-AssessmentRun -Scope $configuration.Assessment.Scope -BaseOutput $configuration.Assessment.OutputRoot -FileSharePath $configuration.Assessment.MappingCsv
    $scriptResults = New-Object 'System.Collections.Generic.List[object]'
    $successfulControls = 0
    $status = 'Success'

    try {
        foreach ($control in $enabledControls) {
            $command = Get-Command -Name $control.ScriptPath -ErrorAction Stop
            $controlStart = Get-Date
            $invokeParameters = @{}

            foreach ($entry in $control.Parameters.GetEnumerator()) {
                $invokeParameters[$entry.Key] = $entry.Value
            }

            if ($command.Parameters.ContainsKey('Run')) {
                $invokeParameters.Run = $run
            }

            if ($command.Parameters.ContainsKey($control.MappingParameter)) {
                $invokeParameters[$control.MappingParameter] = $configuration.Assessment.MappingCsv
            }

            if ($command.Parameters.ContainsKey('Credential') -and $PSBoundParameters.ContainsKey('Credential')) {
                $invokeParameters.Credential = $Credential
            }

            if ($command.Parameters.ContainsKey('Server') -and -not [string]::IsNullOrWhiteSpace($effectiveServer)) {
                $invokeParameters.Server = $effectiveServer
            }

            if ($PSCmdlet.ShouldProcess($control.ScriptPath, "Exécuter le contrôle '$($control.Name)'") ) {
                & $control.ScriptPath @invokeParameters
            }

            $successfulControls++
            $scriptResults.Add([PSCustomObject]@{
                name             = $control.Name
                script_path      = $control.ScriptPath
                status           = 'Success'
                started_at       = $controlStart.ToString('o')
                ended_at         = (Get-Date).ToString('o')
                mapping_parameter = $control.MappingParameter
            }) | Out-Null
        }
    }
    catch {
        $status = 'Failed'
        $scriptResults.Add([PSCustomObject]@{
            name        = if ($control) { $control.Name } else { 'unknown' }
            script_path = if ($control) { $control.ScriptPath } else { '' }
            status      = 'Failed'
            error       = $_.Exception.Message
            ended_at    = (Get-Date).ToString('o')
        }) | Out-Null
        throw
    }
    finally {
        Write-RunManifest -Run $run -Status $status -Scripts $scriptResults.ToArray() -Summary @{
            assessment_name    = $configuration.Assessment.Name
            mapping_csv        = $configuration.Assessment.MappingCsv
            enabled_controls   = $enabledControls.Count
            successful_controls = $successfulControls
        }
        Update-RunsHistory -Run $run -Status $status
    }

    if ($PassThru) {
        return [PSCustomObject]@{
            Run      = $run
            Status   = $status
            Controls = $scriptResults.ToArray()
        }
    }

    return $run
}
