function Import-FileShareMapping {
    <#
    .SYNOPSIS
    Importe et normalise un fichier FileShareMapping.csv au format simplifié.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Path,

        [switch]$Preflight
    )

    $validation = Invoke-FileShareMappingValidation -Path $Path -Preflight:$Preflight
    if (-not $validation.IsValid) {
        $messages = @(
            foreach ($error in $validation.Errors) {
                'Ligne {0}: {1}' -f $error.LineNumber, $error.Message
            }
        )
        throw ($messages -join [Environment]::NewLine)
    }

    return @($validation.Rows)
}
