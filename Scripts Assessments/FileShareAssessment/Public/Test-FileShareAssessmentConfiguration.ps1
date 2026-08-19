function Test-FileShareAssessmentConfiguration {
    <#
    .SYNOPSIS
    Valide une configuration d'assessment.

    .DESCRIPTION
    Retourne un objet simple indiquant si la configuration est valide et, le cas échéant,
    le détail de la première erreur détectée.

    .PARAMETER ConfigurationPath
    Chemin vers le fichier JSON de configuration.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$ConfigurationPath
    )

    try {
        $configuration = Import-FileShareAssessmentConfiguration -ConfigurationPath $ConfigurationPath
        return [PSCustomObject]@{
            IsValid       = $true
            Errors        = @()
            Configuration = $configuration
        }
    }
    catch {
        return [PSCustomObject]@{
            IsValid       = $false
            Errors        = @($_.Exception.Message)
            Configuration = $null
        }
    }
}
