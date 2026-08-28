function Resolve-FileShareSourceIdentifiers {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [array]$Rows,

        [int]$MaxLength = 150
    )

    $grouped = $Rows | Group-Object -Property SourceIdentifierBase
    foreach ($group in $grouped) {
        if ($group.Count -le 1) {
            $group.Group[0].SourceIdentifier = $group.Name
            continue
        }

        foreach ($row in $group.Group) {
            $suffix = Get-FileShareSourceHash -SourcePath $row.SourcePath
            $baseMaxLength = [Math]::Max(1, $MaxLength - ($suffix.Length + 1))
            $base = $row.SourceIdentifierBase
            if ($base.Length -gt $baseMaxLength) {
                $base = $base.Substring(0, $baseMaxLength).TrimEnd('_', '.', ' ')
            }
            if ([string]::IsNullOrWhiteSpace($base)) {
                $base = 'Source'
            }
            $row.SourceIdentifier = '{0}_{1}' -f $base, $suffix
        }
    }

    $remainingCollisions = $Rows | Group-Object -Property SourceIdentifier | Where-Object { $_.Count -gt 1 }
    return [PSCustomObject]@{
        Rows                = $Rows
        RemainingCollisions = @($remainingCollisions)
    }
}
