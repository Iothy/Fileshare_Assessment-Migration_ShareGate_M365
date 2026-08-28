function Resolve-FileShareSourceMetadata {
    <#
    .SYNOPSIS
    Retourne les métadonnées dérivées d'un chemin source local ou UNC.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$SourcePath
    )

    return Get-FileShareSourceMetadata -SourcePath $SourcePath
}
