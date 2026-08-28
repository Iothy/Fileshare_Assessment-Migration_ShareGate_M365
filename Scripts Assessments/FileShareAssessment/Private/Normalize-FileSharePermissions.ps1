function Normalize-FileSharePermissions {
    [CmdletBinding()]
    param(
        [AllowEmptyString()]
        [string]$Value
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        throw 'La colonne Permissions est obligatoire.'
    }

    switch -Regex ($Value.Trim()) {
        '^(?i:yes|oui|true)$' { return 'YES' }
        '^(?i:no|non|false)$' { return 'NO' }
        default {
            throw "La valeur '$Value' n'est pas valide pour Permissions. Utilisez YES/NO (alias Oui/Non/True/False acceptés)."
        }
    }
}
