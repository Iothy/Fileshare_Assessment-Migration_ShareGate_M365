function Import-FileShareAssessmentConfiguration {
    <#
    .SYNOPSIS
    Charge et normalise un fichier de configuration d'assessment.

    .DESCRIPTION
    Lit un fichier JSON, résout les chemins relatifs, valide les sections minimales
    attendues et retourne un objet de configuration prêt à être utilisé par
    Invoke-FileShareAssessment.

    .PARAMETER ConfigurationPath
    Chemin vers le fichier JSON de configuration.

    .EXAMPLE
    Import-FileShareAssessmentConfiguration -ConfigurationPath '.\Config\FileShareAssessment.json'
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$ConfigurationPath
    )

    $resolvedConfigurationPath = Resolve-Path -Path $ConfigurationPath -ErrorAction Stop | Select-Object -ExpandProperty Path
    $configurationDirectory = Split-Path -Path $resolvedConfigurationPath -Parent
    $rawConfiguration = Get-Content -Path $resolvedConfigurationPath -Raw -Encoding UTF8 | ConvertFrom-Json -Depth 20

    if ($null -eq $rawConfiguration.assessment) {
        throw "La section 'assessment' est obligatoire dans $resolvedConfigurationPath."
    }

    if ($null -eq $rawConfiguration.controls) {
        throw "La section 'controls' est obligatoire dans $resolvedConfigurationPath."
    }

    $scope = $rawConfiguration.assessment.scope
    if ([string]::IsNullOrWhiteSpace($scope)) {
        throw "La propriété 'assessment.scope' est obligatoire dans $resolvedConfigurationPath."
    }

    $mappingCsv = Resolve-AssessmentPath -BasePath $configurationDirectory -Path $rawConfiguration.assessment.mappingCsv
    if ([string]::IsNullOrWhiteSpace($mappingCsv)) {
        throw "La propriété 'assessment.mappingCsv' est obligatoire dans $resolvedConfigurationPath."
    }

    if (-not (Test-Path -Path $mappingCsv -PathType Leaf)) {
        throw "Le fichier de mapping '$mappingCsv' est introuvable."
    }

    $outputRoot = Resolve-AssessmentPath -BasePath $configurationDirectory -Path $rawConfiguration.assessment.outputRoot
    if ([string]::IsNullOrWhiteSpace($outputRoot)) {
        throw "La propriété 'assessment.outputRoot' est obligatoire dans $resolvedConfigurationPath."
    }

    $executionOrder = @($rawConfiguration.executionOrder)
    if ($executionOrder.Count -eq 0) {
        $executionOrder = @($rawConfiguration.controls.PSObject.Properties.Name)
    }

    $controls = foreach ($controlName in $executionOrder) {
        $controlDefinition = $rawConfiguration.controls.$controlName
        if ($null -eq $controlDefinition) {
            throw "Le contrôle '$controlName' référencé dans executionOrder est absent de la section 'controls'."
        }

        $scriptPath = Resolve-AssessmentPath -BasePath $configurationDirectory -Path $controlDefinition.script
        if ([string]::IsNullOrWhiteSpace($scriptPath)) {
            throw "Le contrôle '$controlName' doit définir une propriété 'script'."
        }

        if (-not (Test-Path -Path $scriptPath -PathType Leaf)) {
            throw "Le script '$scriptPath' du contrôle '$controlName' est introuvable."
        }

        [PSCustomObject]@{
            Name             = $controlName
            Enabled          = [bool]$controlDefinition.enabled
            Description      = [string]$controlDefinition.description
            ScriptPath       = $scriptPath
            MappingParameter = if ([string]::IsNullOrWhiteSpace($controlDefinition.mappingParameter)) {
                if ($controlName -eq 'report') { 'FileShareMapping' } else { 'MappingCsv' }
            }
            else {
                [string]$controlDefinition.mappingParameter
            }
            Parameters       = if ($null -eq $controlDefinition.parameters) { @{} } else { ConvertTo-AssessmentHashtable -InputObject $controlDefinition.parameters }
        }
    }

    if (@($controls | Where-Object Enabled).Count -eq 0) {
        throw "Aucun contrôle actif n'est défini dans $resolvedConfigurationPath."
    }

    [PSCustomObject]@{
        ConfigurationPath      = $resolvedConfigurationPath
        ConfigurationDirectory = $configurationDirectory
        Assessment             = [PSCustomObject]@{
            Name      = if ([string]::IsNullOrWhiteSpace($rawConfiguration.assessment.name)) { 'FileShareAssessment' } else { [string]$rawConfiguration.assessment.name }
            Scope     = [string]$scope
            MappingCsv = $mappingCsv
            OutputRoot = $outputRoot
            Server    = [string]$rawConfiguration.assessment.server
        }
        ExecutionOrder         = [string[]]$executionOrder
        Controls               = @($controls)
    }
}
