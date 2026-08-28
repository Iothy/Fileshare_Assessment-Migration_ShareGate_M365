function ConvertTo-FileShareSourceIdentifier {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$SourcePath,

        [int]$MaxLength = 150
    )

    $canonical = ConvertTo-CanonicalUncPath -Path $SourcePath
    $parts = @($canonical.TrimStart('\') -split '\\+' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    $rawIdentifier = $parts -join '_'

    $formD = $rawIdentifier.Normalize([Text.NormalizationForm]::FormD)
    $builder = New-Object System.Text.StringBuilder
    foreach ($char in $formD.ToCharArray()) {
        $category = [Globalization.CharUnicodeInfo]::GetUnicodeCategory($char)
        if ($category -eq [Globalization.UnicodeCategory]::NonSpacingMark) {
            continue
        }

        if ($char -match '[A-Za-z0-9._-]') {
            [void]$builder.Append($char)
        }
        elseif ($char -match '\s') {
            # Les espaces sont supprimés (et non remplacés par '_') afin de concaténer les mots
            # d'un même segment, ex. "Sub Subfolder1" -> "SubSubfolder1".
            continue
        }
        elseif ($char -match '[&+/\\]') {
            [void]$builder.Append('_')
        }
        elseif ([int][char]$char -lt 32) {
            continue
        }
        else {
            [void]$builder.Append('_')
        }
    }

    $safe = [regex]::Replace($builder.ToString(), '[<>:"/\\|?*]', '_')
    $safe = [regex]::Replace($safe, '_+', '_')
    $safe = $safe.Trim('_', '.', ' ')
    if ([string]::IsNullOrWhiteSpace($safe)) {
        $safe = 'Source'
    }

    if ($safe.Length -gt $MaxLength) {
        $safe = $safe.Substring(0, $MaxLength).TrimEnd('_', '.', ' ')
    }

    if ([string]::IsNullOrWhiteSpace($safe)) {
        $safe = 'Source'
    }

    return $safe
}
