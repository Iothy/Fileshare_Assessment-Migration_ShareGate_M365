function ConvertTo-CanonicalUncPath {
    [CmdletBinding()]
    param(
        [AllowEmptyString()]
        [string]$Path
    )

    if ([string]::IsNullOrWhiteSpace($Path)) {
        throw 'Le chemin source est obligatoire.'
    }

    $trimmed = $Path.Trim().Replace('/', '\')
    if (-not $trimmed.StartsWith('\\')) {
        throw "Le chemin source '$Path' doit être un chemin UNC commençant par \\ ."
    }

    $segments = @($trimmed.TrimStart('\') -split '\\+' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($segments.Count -lt 2) {
        throw "Le chemin source '$Path' doit contenir au minimum un serveur et un partage."
    }

    return '\\' + ($segments -join '\')
}
