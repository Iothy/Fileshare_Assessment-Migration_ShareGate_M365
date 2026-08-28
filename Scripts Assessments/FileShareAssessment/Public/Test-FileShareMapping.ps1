function Test-FileShareMapping {
    <#
    .SYNOPSIS
    Valide un fichier FileShareMapping.csv au format simplifié.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Path,

        [switch]$Preflight
    )

    return Invoke-FileShareMappingValidation -Path $Path -Preflight:$Preflight
}
