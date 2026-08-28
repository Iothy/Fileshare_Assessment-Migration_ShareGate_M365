function Normalize-FileShareTargetType {
    [CmdletBinding()]
    param(
        [AllowEmptyString()]
        [string]$Value
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        throw 'La colonne TargetType est obligatoire.'
    }

    switch -Regex ($Value.Trim()) {
        '^(?i:sharepoint)$' { return 'SharePoint' }
        '^(?i:onedrive)$'   { return 'OneDrive' }
        default {
            throw "La valeur '$Value' n'est pas supportée pour TargetType. Utilisez SharePoint ou OneDrive."
        }
    }
}
