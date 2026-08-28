function Invoke-FileShareMappingValidation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path,

        [switch]$Preflight
    )

    $errors = New-Object 'System.Collections.Generic.List[object]'
    $warnings = New-Object 'System.Collections.Generic.List[object]'
    $rows = New-Object 'System.Collections.Generic.List[object]'
    $expectedHeaders = Get-FileShareMappingHeaders

    if (-not (Test-Path -Path $Path -PathType Leaf)) {
        $errors.Add([PSCustomObject]@{ LineNumber = 0; Message = "Le fichier de mapping '$Path' est introuvable." }) | Out-Null
        return [PSCustomObject]@{ IsValid = $false; Errors = $errors.ToArray(); Warnings = $warnings.ToArray(); Rows = @() }
    }

    $lines = Get-Content -Path $Path -Encoding UTF8
    if ($lines.Count -eq 0) {
        $errors.Add([PSCustomObject]@{ LineNumber = 1; Message = "Le fichier de mapping '$Path' est vide." }) | Out-Null
        return [PSCustomObject]@{ IsValid = $false; Errors = $errors.ToArray(); Warnings = $warnings.ToArray(); Rows = @() }
    }

    $headerLine = ($lines[0] -replace '^[\uFEFF]+', '').Trim()
    $actualHeaders = @($headerLine -split ';')
    # Décision volontaire : l'import exige les 6 en-têtes exacts et dans le même ordre.
    if ($actualHeaders.Count -ne $expectedHeaders.Count -or (@($actualHeaders) -join '|') -ne (@($expectedHeaders) -join '|')) {
        $errors.Add([PSCustomObject]@{
            LineNumber = 1
            Message    = "En-têtes invalides. Le fichier doit contenir exactement : $($expectedHeaders -join ';')"
        }) | Out-Null
        return [PSCustomObject]@{ IsValid = $false; Errors = $errors.ToArray(); Warnings = $warnings.ToArray(); Rows = @() }
    }

    for ($index = 1; $index -lt $lines.Count; $index++) {
        $lineNumber = $index + 1
        $line = $lines[$index]
        if ([string]::IsNullOrWhiteSpace($line)) {
            continue
        }

        try {
            $row = $line | ConvertFrom-Csv -Delimiter ';' -Header $expectedHeaders | Select-Object -First 1
        }
        catch {
            $errors.Add([PSCustomObject]@{ LineNumber = $lineNumber; Message = "Impossible de lire la ligne $lineNumber du mapping : $($_.Exception.Message)" }) | Out-Null
            continue
        }

        $rawSourcePath = [string]$row.'SourcePath'
        $rawTargetType = [string]$row.'TargetType'
        $rawTargetUrl = [string]$row.'TargetSPOURL'
        $rawTargetFolder = [string]$row.'TargetFolder'
        $rawDateFilter = [string]$row.'DateFilter (YYYY-DD-MM)'
        $rawPermissions = [string]$row.'Permissions'

        $rowErrors = New-Object 'System.Collections.Generic.List[string]'
        $normalizedSourcePath = $null
        $sourceMetadata = $null
        $targetType = $null
        $targetUrl = $null
        $targetFolder = $null
        $dateFilter = $null
        $permissions = $null

        try {
            $normalizedSourcePath = ConvertTo-CanonicalUncPath -Path $rawSourcePath
            $sourceMetadata = Get-FileShareSourceMetadata -SourcePath $normalizedSourcePath
        }
        catch {
            $rowErrors.Add($_.Exception.Message) | Out-Null
        }

        try { $targetType = Normalize-FileShareTargetType -Value $rawTargetType } catch { $rowErrors.Add($_.Exception.Message) | Out-Null }
        try { $targetUrl = Normalize-FileShareTargetSPOUrl -Value $rawTargetUrl } catch { $rowErrors.Add($_.Exception.Message) | Out-Null }
        try { $targetFolder = Normalize-FileShareTargetFolder -Value $rawTargetFolder } catch { $rowErrors.Add($_.Exception.Message) | Out-Null }
        try { $dateFilter = Normalize-FileShareDateFilter -Value $rawDateFilter } catch { $rowErrors.Add($_.Exception.Message) | Out-Null }
        try { $permissions = Normalize-FileSharePermissions -Value $rawPermissions } catch { $rowErrors.Add($_.Exception.Message) | Out-Null }

        foreach ($message in $rowErrors) {
            $errors.Add([PSCustomObject]@{ LineNumber = $lineNumber; Message = $message }) | Out-Null
        }

        if ($rowErrors.Count -gt 0) {
            continue
        }

        $rows.Add(([PSCustomObject]@{
            SourcePath            = $sourceMetadata.SourcePath
            TargetType            = $targetType
            TargetSPOURL          = $targetUrl
            TargetFolder          = $targetFolder
            DateFilter            = $dateFilter
            Permissions           = $permissions
            ServerFqdn            = $sourceMetadata.ServerFqdn
            ServerShort           = $sourceMetadata.ServerShort
            ShareName             = $sourceMetadata.ShareName
            RelativePath          = $sourceMetadata.RelativePath
            LeafName              = $sourceMetadata.LeafName
            SourceIdentifierBase  = $sourceMetadata.SourceIdentifier
            SourceIdentifier      = $sourceMetadata.SourceIdentifier
            LineNumber            = $lineNumber
        })) | Out-Null
    }

    $duplicateSourceGroups = $rows | Group-Object -Property SourcePath | Where-Object { $_.Count -gt 1 }
    foreach ($group in $duplicateSourceGroups) {
        foreach ($duplicateRow in $group.Group) {
            $errors.Add([PSCustomObject]@{
                LineNumber = $duplicateRow.LineNumber
                Message    = "Le chemin source '$($duplicateRow.SourcePath)' est dupliqué dans le mapping après normalisation."
            }) | Out-Null
        }
    }

    if ($rows.Count -gt 0) {
        $identifierResolution = Resolve-FileShareSourceIdentifiers -Rows $rows.ToArray()
        foreach ($collision in $identifierResolution.RemainingCollisions) {
            foreach ($collisionRow in $collision.Group) {
                $errors.Add([PSCustomObject]@{
                    LineNumber = $collisionRow.LineNumber
                    Message    = "Collision d'identifiant non résolue pour '$($collisionRow.SourcePath)' (identifiant : $($collisionRow.SourceIdentifier))."
                }) | Out-Null
            }
        }

        $sortedRows = @($identifierResolution.Rows | Sort-Object SourcePath, LineNumber)
        for ($i = 0; $i -lt $sortedRows.Count; $i++) {
            for ($j = $i + 1; $j -lt $sortedRows.Count; $j++) {
                $left = $sortedRows[$i]
                $right = $sortedRows[$j]
                if ($right.SourcePath.StartsWith($left.SourcePath + '\', [System.StringComparison]::OrdinalIgnoreCase) -or
                    $left.SourcePath.StartsWith($right.SourcePath + '\', [System.StringComparison]::OrdinalIgnoreCase)) {
                    $warnings.Add([PSCustomObject]@{
                        LineNumber = $right.LineNumber
                        Message    = "Chevauchement détecté entre '$($left.SourcePath)' et '$($right.SourcePath)'. Les sous-arborescences risquent d'être analysées plusieurs fois."
                    }) | Out-Null
                }
            }
        }

        $destinationGroups = $sortedRows |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_.TargetSPOURL) } |
            Group-Object { '{0}|{1}' -f $_.TargetSPOURL, $_.TargetFolder }
        foreach ($group in $destinationGroups | Where-Object { $_.Count -gt 1 }) {
            foreach ($destinationRow in $group.Group) {
                $warnings.Add([PSCustomObject]@{
                    LineNumber = $destinationRow.LineNumber
                    Message    = "La destination '$($destinationRow.TargetSPOURL)/$($destinationRow.TargetFolder)' est partagée par plusieurs sources."
                }) | Out-Null
            }
        }

        foreach ($currentRow in $sortedRows) {
            $destinationPath = if ([string]::IsNullOrWhiteSpace($currentRow.TargetFolder)) { $currentRow.TargetSPOURL } else { '{0}/{1}' -f $currentRow.TargetSPOURL, $currentRow.TargetFolder }
            if (-not [string]::IsNullOrWhiteSpace($currentRow.TargetSPOURL) -and $destinationPath.Length -gt 218) {
                $warnings.Add([PSCustomObject]@{
                    LineNumber = $currentRow.LineNumber
                    Message    = "La destination normalisée '$destinationPath' dépasse 218 caractères et peut poser problème côté SharePoint."
                }) | Out-Null
            }

            if ($Preflight) {
                $preflightResult = Test-FileShareSourcePreflight -SourcePath $currentRow.SourcePath
                if (-not $preflightResult.Success) {
                    $warnings.Add([PSCustomObject]@{
                        LineNumber = $currentRow.LineNumber
                        Message    = $preflightResult.Message
                    }) | Out-Null
                }
            }
        }
    }

    return [PSCustomObject]@{
        IsValid  = ($errors.Count -eq 0)
        Errors   = @($errors | Sort-Object LineNumber, Message)
        Warnings = @($warnings | Sort-Object LineNumber, Message -Unique)
        Rows     = @($rows | Sort-Object LineNumber)
    }
}
