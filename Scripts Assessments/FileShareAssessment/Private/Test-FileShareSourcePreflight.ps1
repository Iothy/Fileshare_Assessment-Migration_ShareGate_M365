function Test-FileShareSourcePreflight {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$SourcePath
    )

    try {
        if (-not (Test-Path -Path $SourcePath)) {
            return [PSCustomObject]@{
                Success = $false
                Type    = 'NotFound'
                Message = "Le chemin source '$SourcePath' est introuvable ou inaccessible."
            }
        }

        $null = Get-Item -Path $SourcePath -ErrorAction Stop
        return [PSCustomObject]@{
            Success = $true
            Type    = 'Success'
            Message = ''
        }
    }
    catch [System.UnauthorizedAccessException] {
        return [PSCustomObject]@{
            Success = $false
            Type    = 'AccessDenied'
            Message = "Accès refusé au chemin source '$SourcePath'."
        }
    }
    catch {
        return [PSCustomObject]@{
            Success = $false
            Type    = 'Error'
            Message = "Le chemin source '$SourcePath' n'est pas accessible : $($_.Exception.Message)"
        }
    }
}
