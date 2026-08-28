function Normalize-FileShareDateFilter {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [AllowEmptyString()]
        [string]$Value
    )

    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace($Value)) {
        return ''
    }

    $culture = [System.Globalization.CultureInfo]::InvariantCulture
    $styles = [System.Globalization.DateTimeStyles]::None
    try {
        $parsed = [datetime]::ParseExact($Value.Trim(), 'yyyy-dd-MM', $culture, $styles)
        return $parsed.ToString('yyyy-dd-MM', $culture)
    }
    catch {
        throw "La date '$Value' n'est pas valide. Le format attendu est strictement YYYY-DD-MM (exemple : 2020-31-12)."
    }
}
