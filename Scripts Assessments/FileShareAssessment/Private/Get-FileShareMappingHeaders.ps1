function Get-FileShareMappingHeaders {
    [CmdletBinding()]
    param()

    return @(
        'SourcePath',
        'TargetType',
        'TargetSPOURL',
        'TargetFolder',
        'DateFilter (YYYY-DD-MM)',
        'Permissions'
    )
}
