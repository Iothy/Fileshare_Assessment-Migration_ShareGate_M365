function Normalize-FileShareTargetSPOUrl {
    [CmdletBinding()]
    param(
        [AllowEmptyString()]
        [string]$Value
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        throw 'La colonne TargetSPOURL est obligatoire.'
    }

    $uri = $null
    if (-not [System.Uri]::TryCreate($Value.Trim(), [System.UriKind]::Absolute, [ref]$uri)) {
        throw "L'URL cible '$Value' n'est pas valide."
    }

    if ($uri.Scheme -ne 'https') {
        throw "L'URL cible '$Value' doit utiliser HTTPS."
    }

    $builder = [System.UriBuilder]::new($uri)
    $builder.Path = $builder.Path.TrimEnd('/')
    $builder.Query = ''
    $builder.Fragment = ''
    $normalized = $builder.Uri.AbsoluteUri.TrimEnd('/')
    return $normalized
}
