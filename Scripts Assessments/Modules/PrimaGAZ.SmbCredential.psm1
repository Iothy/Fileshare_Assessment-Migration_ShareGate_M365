<#
.SYNOPSIS
    Module helper pour la gestion de sessions SMB authentifiées sur le FileShare PrimaGAZ.

.DESCRIPTION
    Ce module centralise la gestion propre des connexions SMB avec un compte de service
    dédié (ex: SHVE\SVC_FileShare_Read), dans les cas où la session RDP est ouverte avec
    un compte différent (ex: SHVE\SVC_ShareGate_RDP) qui n'a pas les droits de lecture.

    Windows n'autorise qu'une seule identité SMB par serveur cible. Ce module :
    - Détecte les sessions préexistantes vers le serveur avec un autre compte
    - Établit une session "épinglée" via IPC$ (sans lettre de lecteur)
    - Vérifie l'identité de la session résultante
    - Nettoie proprement en sortie (nettoyage idempotent)

    Fonctions exportées :
    - Connect-PrimaGazFileShare        : Établit la session SMB (idempotente)
    - Disconnect-PrimaGazFileShare     : Supprime uniquement le mapping IPC$ créé par ce module
    - Invoke-WithFileShareCredential   : Wrapper try/finally garantissant le nettoyage
    - Test-FileShareIdentity           : Garde-fou — vérifie l'identité SMB active

.NOTES
    Projet  : PrimaGAZ - Migration FileShare vers M365
    Phase   : 01 - Assessment
    Compat. : PowerShell 5.1+
    Encodage : UTF-8 with BOM
#>

#Requires -Version 5.1

function Connect-PrimaGazFileShare {
    <#
    .SYNOPSIS
        Établit une session SMB authentifiée vers un serveur de fichiers.

    .DESCRIPTION
        Avant tout accès :
        1. Vérifie si une session SMB valide (même compte) existe déjà → retour immédiat (idempotent).
        2. Si une session avec un AUTRE compte existe → échoue sans modifier la session.
        3. Établit une nouvelle session via un mapping IPC$ éphémère (Persistent:$false).
        4. Vérifie l'identité de la session résultante via Get-SmbConnection.
        5. Purge le mot de passe en clair immédiatement après usage.

    .PARAMETER Server
        Nom DNS ou IP du serveur SMB cible.

    .PARAMETER Credential
        PSCredential à utiliser pour la connexion SMB.

    .OUTPUTS
        PSCustomObject avec propriétés :
          Server           : nom du serveur
          ConnectedAs      : UserName de la session établie
          Timestamp        : horodatage de la connexion
          AlreadyConnected : $true si la session existait déjà (idempotent)

    .EXAMPLE
        $cred = Get-Credential -UserName 'SHVE\SVC_FileShare_Read'
        Connect-PrimaGazFileShare -Server 'ntx-pa-fs01.primagaz.fr' -Credential $cred
    #>
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Server,

        [Parameter(Mandatory = $true)]
        [System.Management.Automation.PSCredential]
        [System.Management.Automation.Credential()]
        $Credential
    )

    $targetUNC = "\\$Server"
    $ipcUNC    = "\\$Server\IPC$"

    # ──────────────────────────────────────────────────────────────
    # 1. Idempotence : la session est-elle déjà bonne ?
    # ──────────────────────────────────────────────────────────────
    $existingConnections = @()
    try {
        $existingConnections = @(Get-SmbConnection -ErrorAction SilentlyContinue |
            Where-Object { $_.ServerName -eq $Server })
    }
    catch {
        # Get-SmbConnection peut être absent ou non autorisé — ignorer
    }

    if ($existingConnections.Count -gt 0) {
        $alreadyGood = $existingConnections | Where-Object {
            $_.UserName -like "*$($Credential.UserName)*" -or
            $Credential.UserName -like "*$($_.UserName)*"
        }
        if ($alreadyGood) {
            Write-Verbose "[SmbCredential] Session SMB vers $Server déjà active avec '$($Credential.UserName)' — rien à faire."
            return [PSCustomObject]@{
                Server           = $Server
                ConnectedAs      = $Credential.UserName
                Timestamp        = Get-Date
                AlreadyConnected = $true
            }
        }

        throw "[SmbCredential] Une session SMB existe déjà vers '$Server' avec '$($existingConnections[0].UserName)'. Elle n'a pas été modifiée. Fermez-la explicitement ou exécutez l'assessment sous l'identité déjà connectée."
    }

    # ──────────────────────────────────────────────────────────────
    # 2. Établir la session via IPC$
    # ──────────────────────────────────────────────────────────────
    $plainPassword = $null
    try {
        $plainPassword = [System.Net.NetworkCredential]::new('', $Credential.Password).Password

        $smbParams = @{
            RemotePath  = $ipcUNC
            UserName    = $Credential.UserName
            Password    = $plainPassword
            Persistent  = $false
            ErrorAction = 'Stop'
        }

        New-SmbMapping @smbParams | Out-Null
        Write-Verbose "[SmbCredential] Mapping IPC$ créé vers $ipcUNC avec '$($Credential.UserName)'."
    }
    catch {
        throw "[SmbCredential] Impossible d'établir la session SMB vers $ipcUNC avec '$($Credential.UserName)' : $($_.Exception.Message)"
    }
    finally {
        # Purge immédiate du mot de passe en clair
        if ($null -ne $plainPassword) {
            Remove-Variable -Name plainPassword -Force -ErrorAction SilentlyContinue
        }
    }

    # ──────────────────────────────────────────────────────────────
    # 3. Vérification de l'identité résultante
    # ──────────────────────────────────────────────────────────────
    Start-Sleep -Milliseconds 200   # laisser le temps au mapping de s'enregistrer

    $verifiedConnection = $null
    try {
        $verifiedConnection = Get-SmbConnection -ErrorAction SilentlyContinue |
            Where-Object { $_.ServerName -eq $Server } |
            Select-Object -First 1
    }
    catch {
        Write-Verbose "[SmbCredential] Get-SmbConnection indisponible pour vérification — poursuite sans garantie."
    }

    if ($verifiedConnection) {
        # Comparaison insensible à la casse et à la forme DOMAIN\user vs user@domain
        $expectedUser = $Credential.UserName
        $actualUser   = $verifiedConnection.UserName

        $match = ($actualUser -eq $expectedUser) -or
                 ($actualUser -like "*\$($expectedUser.Split('\')[-1])*") -or
                 ($expectedUser -like "*\$($actualUser.Split('\')[-1])*")

        if (-not $match) {
            throw "[SmbCredential] Vérification identité ÉCHOUÉE : attendu '$expectedUser', session SMB retourne '$actualUser'. La connexion au serveur utilise peut-être encore un cache Kerberos antérieur."
        }
        Write-Verbose "[SmbCredential] Identité SMB vérifiée : '$actualUser'."
    }
    else {
        Write-Verbose "[SmbCredential] Impossible de vérifier l'identité via Get-SmbConnection (cmdlet indisponible ou connexion non encore visible)."
    }

    return [PSCustomObject]@{
        Server           = $Server
        ConnectedAs      = $Credential.UserName
        Timestamp        = Get-Date
        AlreadyConnected = $false
    }
}

function Disconnect-PrimaGazFileShare {
    <#
    .SYNOPSIS
        Supprime tous les mappings SMB créés vers un serveur de fichiers.

    .DESCRIPTION
        Supprime le mapping IPC$ et tout autre mapping vers le serveur cible.
        Idempotent : ne lève pas d'erreur si aucun mapping n'existe.

    .PARAMETER Server
        Nom DNS ou IP du serveur SMB cible.

    .EXAMPLE
        Disconnect-PrimaGazFileShare -Server 'ntx-pa-fs01.primagaz.fr'
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Server
    )

    $targetUNC = "\\$Server"
    $ipcUNC    = "\\$Server\IPC$"
    $removed   = 0

    # Suppression du mapping IPC$ explicite
    try {
        Remove-SmbMapping -RemotePath $ipcUNC -Force -ErrorAction SilentlyContinue
        $removed++
    }
    catch {
        Write-Verbose "[SmbCredential] Remove-SmbMapping IPC$ : $($_.Exception.Message)"
    }

    # Nettoyage complémentaire du seul IPC$ (pour les connexions non visibles via Get-SmbMapping)
    try {
    & net use "$targetUNC\IPC$" /delete /y 2>&1 | Out-Null
    }
    catch {
        Write-Verbose "[SmbCredential] net use IPC$ /delete : $($_.Exception.Message)"
    }

    Write-Verbose "[SmbCredential] Déconnexion SMB de '$Server' terminée ($removed mapping(s) supprimé(s))."
}

function Invoke-WithFileShareCredential {
    <#
    .SYNOPSIS
        Exécute un ScriptBlock sous une session SMB authentifiée, avec nettoyage garanti.

    .DESCRIPTION
        Wrapper try/finally qui :
        1. Appelle Connect-PrimaGazFileShare
        2. Exécute le ScriptBlock fourni
        3. Appelle Disconnect-PrimaGazFileShare dans finally (garanti même en cas d'exception ou Ctrl+C)

    .PARAMETER Server
        Nom DNS ou IP du serveur SMB cible.

    .PARAMETER Credential
        PSCredential à utiliser pour la connexion SMB.

    .PARAMETER ScriptBlock
        Bloc de code à exécuter sous la session SMB.

    .EXAMPLE
        $cred = Get-Credential -UserName 'SHVE\SVC_FileShare_Read'
        Invoke-WithFileShareCredential -Server 'ntx-pa-fs01.primagaz.fr' -Credential $cred -ScriptBlock {
            Get-ChildItem '\\ntx-pa-fs01.primagaz.fr\EQUIPE'
        }
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Server,

        [Parameter(Mandatory = $true)]
        [System.Management.Automation.PSCredential]
        [System.Management.Automation.Credential()]
        $Credential,

        [Parameter(Mandatory = $true)]
        [scriptblock]$ScriptBlock
    )

    $connected = $false
    try {
        Connect-PrimaGazFileShare -Server $Server -Credential $Credential | Out-Null
        $connected = $true
        & $ScriptBlock
    }
    finally {
        if ($connected) {
            Disconnect-PrimaGazFileShare -Server $Server
        }
    }
}

function Test-FileShareIdentity {
    <#
    .SYNOPSIS
        Vérifie que la session SMB active vers un serveur utilise bien le compte attendu.

    .DESCRIPTION
        Garde-fou à appeler au début de chaque script d'assessment pour s'assurer que
        les accès réseau utilisent bien le compte de service dédié et non le compte RDP.

        Lève une erreur terminating si l'identité ne correspond pas.
        Passe sans erreur si Get-SmbConnection est indisponible (degraded mode).

    .PARAMETER Server
        Nom DNS ou IP du serveur SMB cible.

    .PARAMETER ExpectedUserName
        Nom d'utilisateur attendu (ex: 'SHVE\SVC_FileShare_Read').

    .EXAMPLE
        Test-FileShareIdentity -Server 'ntx-pa-fs01.primagaz.fr' -ExpectedUserName 'SHVE\SVC_FileShare_Read'
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Server,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$ExpectedUserName
    )

    $connection = $null
    try {
        $connection = Get-SmbConnection -ErrorAction SilentlyContinue |
            Where-Object { $_.ServerName -eq $Server } |
            Select-Object -First 1
    }
    catch {
        Write-Verbose "[SmbCredential] Get-SmbConnection indisponible — vérification identité ignorée."
        return
    }

    if (-not $connection) {
        Write-Verbose "[SmbCredential] Aucune session SMB active vers '$Server' — vérification ignorée (peut être normal avant le premier accès)."
        return
    }

    $actualUser = $connection.UserName

    $match = ($actualUser -eq $ExpectedUserName) -or
             ($actualUser -like "*\$($ExpectedUserName.Split('\')[-1])*") -or
             ($ExpectedUserName -like "*\$($actualUser.Split('\')[-1])*")

    if (-not $match) {
        throw "[SmbCredential] ERREUR IDENTITÉ : la session SMB vers '$Server' utilise '$actualUser' au lieu de '$ExpectedUserName'. Appelez Connect-PrimaGazFileShare d'abord."
    }

    Write-Verbose "[SmbCredential] Identité SMB OK : '$actualUser' correspond à '$ExpectedUserName'."
}

Export-ModuleMember -Function Connect-PrimaGazFileShare, Disconnect-PrimaGazFileShare, Invoke-WithFileShareCredential, Test-FileShareIdentity
