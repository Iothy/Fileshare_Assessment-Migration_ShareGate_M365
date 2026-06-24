<#
.SYNOPSIS
    Supprime les dossiers vides migres par ShareGate, sous l'arborescence cible
    definie dans un fichier de mapping (CSV), pour des cibles SharePoint Online,
    canaux Teams standard ("General") et canaux Teams prives.

.DESCRIPTION
    Etape finale d'une migration FileShare -> M365.

    Pour chaque ligne du mapping :
      - Connexion PnP au site TargetSPOURL.
      - Resolution de la racine de nettoyage en fonction de TargetType :
          * SharePoint            -> bibliotheque par defaut du site
          * Teams-Channel General -> <lib>/General/<...>
          * Teams-Private-Channel -> bibliotheque par defaut du site dedie au canal prive
        La bibliotheque par defaut est detectee dynamiquement (BaseTemplate 101),
        ce qui gere "Shared Documents" vs "Documents" et les sites multilingues.
      - Parcours recursif (post-ordre) du dossier racine resolu.
      - Tout sous-dossier ne contenant aucun fichier (ni direct, ni dans ses
        sous-dossiers) est supprime.
      - Le dossier racine TargetFolder n'est PAS supprime meme s'il devient vide.
      - Garde-fous : refus de supprimer la racine de bibliotheque ou un dossier
        "General" (racine de canal Teams).

    Le mode -WhatIf permet une simulation sans suppression. Un rapport CSV du
    bilan est exporte dans C:\migrationFactory\output.

.PARAMETER BatchFile
    Chemin vers le CSV de mapping (ou nom de fichier present dans
    C:\migrationFactory\input).
    Colonnes attendues : TargetType, TargetSPOURL, TargetFolder.
    Colonnes ignorees mais tolerees : SourcePath, DateFilter (YYYY-MM-DD), Permissions.

.PARAMETER WhatIf
    Si specifie, log uniquement ce qui serait supprime sans effectuer de
    suppression.

.EXAMPLE
    .\05-Remove-EmptySPOFolders.ps1 -BatchFile .\SPO-Input_ALL_20260622.csv -WhatIf

.EXAMPLE
    .\05-Remove-EmptySPOFolders.ps1 -BatchFile SPO-Input_ALL_20260622.csv
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true)]
    [string]$BatchFile
)

# Racine de la migration factory
$FactoryRoot = 'C:\migrationFactory'

# Constantes
$ClientId         = '0f78653b-2b30-47f2-9d09-3c17709f118a'
$CredentialTarget = 'ShareGate3'
$Delimiter        = ';'

# Dossiers racine a NE JAMAIS supprimer (racine de canal Teams standard, etc.)
$ProtectedLeafNames = @('General')

# --- 0. Pre-requis ---------------------------------------------------------
Import-Module PnP.PowerShell    -ErrorAction Stop
Import-Module CredentialManager -ErrorAction Stop

# Structure migrationFactory
$InputDir  = Join-Path $FactoryRoot 'input'
$OutputDir = Join-Path $FactoryRoot 'output'
$LogsDir   = Join-Path $FactoryRoot 'logs'
foreach ($d in $InputDir,$OutputDir,$LogsDir) {
    if (-not (Test-Path -LiteralPath $d)) { New-Item -ItemType Directory -Path $d -Force | Out-Null }
}

# Resolution du fichier batch
if (-not (Test-Path -LiteralPath $BatchFile)) {
    $candidate = Join-Path $InputDir (Split-Path -Leaf $BatchFile)
    if (Test-Path -LiteralPath $candidate) {
        $BatchFile = $candidate
    } else {
        throw "Fichier batch introuvable : '$BatchFile' (cherche aussi dans '$InputDir')."
    }
}
$BatchFile = (Resolve-Path -LiteralPath $BatchFile).Path
$batchBase = [System.IO.Path]::GetFileNameWithoutExtension($BatchFile)

# Credential
$cred = Get-StoredCredential -Target $CredentialTarget
if (-not $cred) {
    throw "Credential '$CredentialTarget' introuvable dans le Credential Manager."
}

# Log mensuel : PrimagaZ_Remove_EmptyFolder_M365-XX.log
$monthTag = Get-Date -Format 'MM'
$execLog  = Join-Path $LogsDir ("PrimagaZ_Remove_EmptyFolder_M365-{0}.log" -f $monthTag)

# Rapport CSV de bilan (tracabilite client)
$stamp     = Get-Date -Format 'yyyyMMdd_HHmmss'
$reportCsv = Join-Path $OutputDir ("Remove_EmptyFolder_Report_{0}_{1}.csv" -f $batchBase, $stamp)

# --- Logging ---------------------------------------------------------------
function Write-Log {
    param($Name, $Msg, $Type = "INFO", $Logpath)
    $ts    = Get-Date -Format "yyyyMMdd_HHmmss"
    $line  = "$ts ; $Type ; $Name ; $Msg"
    $color = switch ($Type) { "ERROR" { "Red" } "WARN" { "Yellow" } default { "Green" } }
    Write-Host $line -ForegroundColor $color
    if ($Logpath) {
        try {
            $fs = [System.IO.File]::Open($Logpath, [System.IO.FileMode]::Append, [System.IO.FileAccess]::Write, [System.IO.FileShare]::ReadWrite)
            $sw = [System.IO.StreamWriter]::new($fs, [System.Text.Encoding]::UTF8)
            $sw.WriteLine($line); $sw.Flush(); $sw.Close(); $fs.Close()
        } catch {
            Start-Sleep -Milliseconds (Get-Random -Minimum 50 -Maximum 200)
            $line | Out-File -FilePath $Logpath -Append -Encoding UTF8 -ErrorAction SilentlyContinue
        }
    }
}
$log_params = @{ Name = "05-Remove-EmptySPOFolders" ; Logpath = $execLog }

$simulate = [bool]$WhatIfPreference

Write-Log @log_params -Msg "=== Demarrage 05-Remove-EmptySPOFolders ===" -Type INFO
Write-Log @log_params -Msg "Batch       : $BatchFile"        -Type INFO
Write-Log @log_params -Msg "ClientId    : $ClientId"         -Type INFO
Write-Log @log_params -Msg "Compte      : $($cred.UserName)" -Type INFO
Write-Log @log_params -Msg "WhatIf      : $simulate"         -Type INFO
Write-Log @log_params -Msg "Log         : $execLog"          -Type INFO
Write-Log @log_params -Msg "Rapport     : $reportCsv"        -Type INFO

# --- 1. Lecture CSV --------------------------------------------------------
$rows = Import-Csv -Path $BatchFile -Delimiter $Delimiter

foreach ($col in 'TargetType','TargetSPOURL','TargetFolder') {
    if (-not ($rows | Get-Member -Name $col -MemberType NoteProperty)) {
        throw "Colonne '$col' absente de '$BatchFile'."
    }
}

# --- 2. Resolution de la racine selon TargetType ---------------------------
# Retourne un objet : ServerRelative (chemin server-relative du dossier racine
# a nettoyer), LibRoot (racine de bibliotheque server-relative), Leaf.
function Resolve-TargetRoot {
    param(
        [Parameter(Mandatory)] [string]$TargetType,
        [Parameter(Mandatory)] [string]$TargetFolder,
        [Parameter(Mandatory)] $Connection
    )

    # Normalise les separateurs (\ -> /), supprime les slashes superflus
    $folder = ($TargetFolder -replace '\\','/' -replace '/+','/').Trim('/')

    # Detection dynamique de la bibliotheque par defaut (gere "Documents" vs
    # "Shared Documents" et les sites multilingues).
    $defaultLib = Get-PnPList -Connection $Connection -ErrorAction Stop |
                  Where-Object { $_.BaseTemplate -eq 101 -and -not $_.Hidden } |
                  Sort-Object -Property @{ E = { $_.RootFolder.ServerRelativeUrl.Length } } |
                  Select-Object -First 1
    if (-not $defaultLib) {
        throw "Aucune bibliotheque de documents (BaseTemplate 101) trouvee sur le site."
    }
    $libRoot = $defaultLib.RootFolder.ServerRelativeUrl.TrimEnd('/')  # ex: /sites/xxx/Shared Documents

    switch -Wildcard ($TargetType.Trim()) {

        'SharePoint' {
            # Le TargetFolder contient DEJA "Shared Documents/..." ou "Documents/..."
            # -> on retire le prefixe lib s'il est present pour eviter le doublon.
            $clean   = $folder -replace '^(Shared Documents|Documents)/', ''
            $server  = "$libRoot/$clean"
        }

        'Teams-Channel*' {
            # Canal standard : dossier <Channel>/... DANS la lib du site d'equipe.
            $clean  = $folder -replace '^(Shared Documents|Documents)/', ''
            $server = "$libRoot/$clean"
        }

        'Teams-Private-Channel' {
            # Canal prive : site collection dediee (URL deja specifique).
            $clean  = $folder -replace '^(Shared Documents|Documents)/', ''
            $server = "$libRoot/$clean"
        }

        default {
            throw "TargetType inconnu : '$TargetType'"
        }
    }

    $server = ($server -replace '/+','/').TrimEnd('/')
    [pscustomobject]@{
        ServerRelative = $server
        LibRoot        = $libRoot
        Leaf           = ($server -split '/')[-1]
    }
}

# --- 3. Fonction recursive : suppression des sous-dossiers vides -----------
# Un SEUL appel Get-PnPFolderItem par dossier (fichiers + dossiers), puis tri
# local. Decremente le compteur d'enfants au fil des suppressions plutot que de
# relire le serveur.
function Remove-EmptySubFolders {
    param(
        [Parameter(Mandatory)] $Connection,
        [Parameter(Mandatory)] [string]$FolderServerRelativeUrl,
        [Parameter(Mandatory)] [bool]$IsRoot,
        [Parameter(Mandatory)] [bool]$Simulate,
        [Parameter(Mandatory)] [string]$SiteUrl
    )

    $deletedCount = 0

    # Lister TOUS les items (un seul appel reseau)
    $items = @()
    try {
        $items = Get-PnPFolderItem -FolderSiteRelativeUrl $FolderServerRelativeUrl `
                                   -Connection $Connection `
                                   -ErrorAction Stop |
                 Where-Object { $_.Name -and $_.Name -ne 'Forms' }
    } catch {
        Write-Log @log_params -Msg "[WARN] $SiteUrl  ->  $FolderServerRelativeUrl : enumeration impossible ($($_.Exception.Message))" -Type WARN
        return 0
    }

    # Separer fichiers vs dossiers localement
    $folders = @($items | Where-Object { $_.GetType().Name -eq 'Folder' })
    $files   = @($items | Where-Object { $_.GetType().Name -eq 'File' })

    $childFolderCount = $folders.Count

    foreach ($child in $folders) {
        $childPath = "$FolderServerRelativeUrl/$($child.Name)"
        $childDeleted = Remove-EmptySubFolders -Connection $Connection `
                                               -FolderServerRelativeUrl $childPath `
                                               -IsRoot $false `
                                               -Simulate $Simulate `
                                               -SiteUrl $SiteUrl
        $deletedCount += $childDeleted

        # Si l'enfant direct a ete supprime, on le retire du compte local.
        if ($childDeleted -gt 0 -and (Test-ChildDeleted -Connection $Connection -ChildPath $childPath)) {
            $childFolderCount--
        }
    }

    if ($IsRoot) {
        # On ne supprime jamais la racine TargetFolder.
        return $deletedCount
    }

    # Le dossier courant est vide si : aucun fichier ET plus aucun sous-dossier.
    if ($files.Count -eq 0 -and $childFolderCount -le 0) {
        if ($Simulate) {
            Write-Log @log_params -Msg "[WHATIF] $SiteUrl  ->  $FolderServerRelativeUrl serait supprime" -Type INFO
            $deletedCount++
        } else {
            $leaf   = ($FolderServerRelativeUrl -split '/')[-1]
            $parent = $FolderServerRelativeUrl.Substring(0, $FolderServerRelativeUrl.LastIndexOf('/'))
            try {
                Remove-PnPFolder -Name $leaf `
                                 -Folder $parent `
                                 -Connection $Connection `
                                 -Force `
                                 -ErrorAction Stop
                Write-Log @log_params -Msg "[DEL] $SiteUrl  ->  $FolderServerRelativeUrl supprime" -Type INFO
                $deletedCount++
            } catch {
                Write-Log @log_params -Msg "[KO] $SiteUrl  ->  $FolderServerRelativeUrl : $($_.Exception.Message)" -Type ERROR
            }
        }
    }

    return $deletedCount
}

# Verifie qu'un dossier n'existe plus (apres tentative de suppression).
function Test-ChildDeleted {
    param(
        [Parameter(Mandatory)] $Connection,
        [Parameter(Mandatory)] [string]$ChildPath
    )
    $exists = Get-PnPFolder -Url $ChildPath -Connection $Connection -ErrorAction SilentlyContinue
    return (-not $exists)
}

# --- 4. Boucle principale --------------------------------------------------
# Dedoublonner sur (TargetType + TargetSPOURL + TargetFolder) pour ne traiter
# qu'une fois chaque arborescence racine.
$targets = $rows |
    Where-Object { "$($_.TargetSPOURL)".Trim() -and "$($_.TargetFolder)".Trim() -and "$($_.TargetType)".Trim() } |
    Select-Object @{N='Type';  E={ "$($_.TargetType)".Trim() }},
                  @{N='Url';   E={ "$($_.TargetSPOURL)".Trim() }},
                  @{N='Folder';E={ "$($_.TargetFolder)".Trim() }} |
    Sort-Object Type, Url, Folder -Unique

Write-Log @log_params -Msg "Cibles uniques a parcourir : $($targets.Count)" -Type INFO

$results = foreach ($t in $targets) {

    $conn    = $null
    $deleted = 0

    try {
        $conn = Connect-PnPOnline -Url $t.Url `
                                  -ClientId $ClientId `
                                  -Credentials $cred `
                                  -ReturnConnection `
                                  -ErrorAction Stop

        # Resolution de la racine selon le type de cible
        $root = Resolve-TargetRoot -TargetType $t.Type -TargetFolder $t.Folder -Connection $conn

        # --- Garde-fous de securite ---
        if ($root.ServerRelative -eq $root.LibRoot) {
            Write-Log @log_params -Msg "[SKIP-SECURITE] $($t.Url)  ->  racine de bibliotheque protegee ($($root.ServerRelative))" -Type WARN
            [pscustomobject]@{ Type=$t.Type; Url=$t.Url; Folder=$t.Folder; Root=$root.ServerRelative; Deleted=0; Status='Skipped'; Message='Racine de bibliotheque protegee' }
            continue
        }
        if ($ProtectedLeafNames -contains $root.Leaf) {
            Write-Log @log_params -Msg "[SKIP-SECURITE] $($t.Url)  ->  dossier protege '$($root.Leaf)' ($($root.ServerRelative))" -Type WARN
            [pscustomobject]@{ Type=$t.Type; Url=$t.Url; Folder=$t.Folder; Root=$root.ServerRelative; Deleted=0; Status='Skipped'; Message="Dossier protege '$($root.Leaf)'" }
            continue
        }

        # Verifier l'existence de la racine
        $rootExists = Get-PnPFolder -Url $root.ServerRelative -Connection $conn -ErrorAction SilentlyContinue
        if (-not $rootExists) {
            Write-Log @log_params -Msg "[SKIP] $($t.Url)  ->  $($root.ServerRelative) inexistant" -Type WARN
            [pscustomobject]@{ Type=$t.Type; Url=$t.Url; Folder=$t.Folder; Root=$root.ServerRelative; Deleted=0; Status='Skipped'; Message='Racine inexistante' }
            continue
        }

        Write-Log @log_params -Msg "[..] $($t.Url)  ->  parcours de $($root.ServerRelative)" -Type INFO

        $deleted = Remove-EmptySubFolders -Connection $conn `
                                          -FolderServerRelativeUrl $root.ServerRelative `
                                          -IsRoot $true `
                                          -Simulate $simulate `
                                          -SiteUrl $t.Url

        Write-Log @log_params -Msg "[OK] $($t.Url)  ->  $($root.ServerRelative) : $deleted dossier(s) vide(s)" -Type INFO

        [pscustomobject]@{
            Type    = $t.Type
            Url     = $t.Url
            Folder  = $t.Folder
            Root    = $root.ServerRelative
            Deleted = $deleted
            Status  = 'Done'
            Message = ''
        }
    }
    catch {
        Write-Log @log_params -Msg "[KO] $($t.Url)  ->  $($_.Exception.Message)" -Type ERROR
        [pscustomobject]@{ Type=$t.Type; Url=$t.Url; Folder=$t.Folder; Root=''; Deleted=$deleted; Status='Failed'; Message=$_.Exception.Message }
    }
    finally {
        if ($conn) {
            try { Disconnect-PnPOnline -Connection $conn -ErrorAction SilentlyContinue } catch {}
        }
    }
}

# --- 5. Export rapport + bilan ---------------------------------------------
try {
    $results | Export-Csv -Path $reportCsv -Delimiter $Delimiter -NoTypeInformation -Encoding UTF8
    Write-Log @log_params -Msg "Rapport exporte : $reportCsv" -Type INFO
} catch {
    Write-Log @log_params -Msg "[WARN] Export rapport impossible : $($_.Exception.Message)" -Type WARN
}

$totalDeleted = ($results | Measure-Object -Property Deleted -Sum).Sum
$done    = ($results | Where-Object Status -eq 'Done').Count
$failed  = ($results | Where-Object Status -eq 'Failed').Count
$skipped = ($results | Where-Object Status -eq 'Skipped').Count

if ($simulate) {
    Write-Log @log_params -Msg "Bilan (WhatIf) : $totalDeleted dossier(s) seraient supprime(s) / $done OK / $failed KO / $skipped ignore(s)" -Type INFO
} else {
    Write-Log @log_params -Msg "Bilan : $totalDeleted dossier(s) supprime(s) / $done OK / $failed KO / $skipped ignore(s)" -Type INFO
}
Write-Log @log_params -Msg "Log mensuel : $execLog" -Type INFO
Write-Log @log_params -Msg "=== Fin 05-Remove-EmptySPOFolders ===" -Type INFO

if ($failed -gt 0) { exit 1 } else { exit 0 }
