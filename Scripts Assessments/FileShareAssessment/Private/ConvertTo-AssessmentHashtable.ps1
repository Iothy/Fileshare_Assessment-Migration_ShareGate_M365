function ConvertTo-AssessmentHashtable {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $true)]
        $InputObject
    )

    process {
        if ($null -eq $InputObject) {
            return @{}
        }

        if ($InputObject -is [hashtable]) {
            return $InputObject
        }

        $table = @{}
        foreach ($property in $InputObject.PSObject.Properties) {
            $value = $property.Value
            if ($null -eq $value) {
                $table[$property.Name] = $null
            }
            elseif ($value -is [pscustomobject] -or $value -is [hashtable]) {
                $table[$property.Name] = ConvertTo-AssessmentHashtable -InputObject $value
            }
            elseif ($value -is [System.Collections.IEnumerable] -and $value -isnot [string]) {
                $table[$property.Name] = @(
                    foreach ($item in $value) {
                        if ($item -is [pscustomobject] -or $item -is [hashtable]) {
                            ConvertTo-AssessmentHashtable -InputObject $item
                        }
                        else {
                            $item
                        }
                    }
                )
            }
            else {
                $table[$property.Name] = $value
            }
        }

        return $table
    }
}
