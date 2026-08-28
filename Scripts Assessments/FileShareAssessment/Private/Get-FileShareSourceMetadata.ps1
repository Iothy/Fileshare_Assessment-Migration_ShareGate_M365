function Get-FileShareSourceMetadata {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$SourcePath
    )

    $canonical = ConvertTo-CanonicalUncPath -Path $SourcePath
    $segments = @($canonical.TrimStart('\') -split '\\+' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($segments.Count -lt 2) {
        throw "Le chemin source '$SourcePath' doit contenir un serveur et un partage."
    }

    $serverFqdn = $segments[0]
    $serverShort = ($serverFqdn -split '\.')[0]
    $shareName = $segments[1]
    $relativeSegments = [string[]]@()
    if ($segments.Count -gt 2) {
        $relativeSegments = [string[]]($segments[2..($segments.Count - 1)])
    }
    $relativePath = ($relativeSegments -join '/')
    $leafName = if ($relativeSegments.Length -gt 0) { $relativeSegments[$relativeSegments.Length - 1] } else { $shareName }

    [PSCustomObject]@{
        SourcePath       = $canonical
        ServerFqdn       = $serverFqdn
        ServerShort      = $serverShort
        ShareName        = $shareName
        RelativePath     = $relativePath
        LeafName         = $leafName
        SourceIdentifier = ConvertTo-FileShareSourceIdentifier -SourcePath $canonical
    }
}
