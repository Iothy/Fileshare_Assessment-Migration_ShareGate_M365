function Normalize-FileShareTargetFolder {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [AllowEmptyString()]
        [string]$Value
    )

    if ($null -eq $Value) {
        return ''
    }

    $trimmed = $Value.Trim()
    if ([string]::IsNullOrWhiteSpace($trimmed)) {
        return ''
    }

    if ($trimmed -match '[\x00-\x1F]') {
        throw 'TargetFolder contient des caractères de contrôle non autorisés.'
    }

    $normalized = $trimmed.Replace('\', '/')
    $normalized = [regex]::Replace($normalized, '/+', '/')
    $normalized = $normalized.Trim('/')
    return $normalized
}
