function Get-FileShareSourceHash {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$SourcePath,

        [int]$Length = 6
    )

    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $bytes = [System.Text.Encoding]::UTF8.GetBytes((ConvertTo-CanonicalUncPath -Path $SourcePath))
        $hashBytes = $sha.ComputeHash($bytes)
    }
    finally {
        $sha.Dispose()
    }

    $hash = -join ($hashBytes | ForEach-Object { $_.ToString('x2') })
    return $hash.Substring(0, [Math]::Min($Length, $hash.Length))
}
