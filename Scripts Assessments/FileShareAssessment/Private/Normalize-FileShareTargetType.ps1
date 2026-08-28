function Normalize-FileShareTargetType {
    [CmdletBinding()]
    param(
        [AllowEmptyString()]
        [string]$Value
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return ''
    }

    switch -Regex ($Value.Trim()) {
        '^(?i:sharepoint)$' { return 'SharePoint' }
        '^(?i:onedrive)$'   { return 'OneDrive' }
        '^(?i:teams-channel general)$' { return 'Teams-Channel General' }
        '^(?i:teams-private-channel)$' { return 'Teams-Private-Channel' }
        default {
            throw "La valeur '$Value' n'est pas supportée pour TargetType. Utilisez SharePoint, OneDrive, Teams-Channel General ou Teams-Private-Channel."
        }
    }
}
