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
        La bibliotheque par defaut est detectee dynamiquement (BaseTemplate 101,
        RootFolder.Name = "Shared Documents"/"Documents"), en EXCLUANT les
        bibliotheques systeme (SiteAssets, Style Library, Form Templates, etc.).
        Cela gere aussi "Shared Documents" vs "Documents" et les sites multilingues.
      - Parcours recursif (post-ordre) du dossier racine resolu.
      - Tout sous-dossier ne contenant aucun fichier (ni direct, ni dans ses
        sous-dossiers) est supprime.
      - Le dossier racine TargetFolder n'est PAS supprime meme s'il devient vide.
      - Garde-fous : refus de supprimer la racine de bibliotheque ou un dossier
        "General" (racine de canal Teams).

    IMPORTANT - Gestion des chemins :
      - Get-PnPFolderItem  attend un chemin SITE-relative (ex: "Shared Documents/X").
      - Remove-PnPFolder   attend un chemin SERVER-relative (ex: "/sites/web/Shared Documents").
      Le script manipule donc en interne le chemin SITE-relative pour l'enumeration
      et derive le SERVER-relative (via WebServerRelativeUrl) uniquement pour la
      suppression.

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

# Bibliotheques systeme a EXCLURE de la detection (meme BaseTemplate 101).
# Comparaison faite sur RootFolder.Name (nom interne, stable).
$ExcludedLibInternalNames = @(
    'SiteAssets',
    'Style Library',
    'FormServerTemplates',
    'Form Templates',
    'SitePages',
    'Site Pages',
    'Lists',
    'Preservation Hold Library'
)

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

# --- 2. Detection de la bibliotheque par defaut ----------------------------
# SharePoint n'expose pas de propriete "bibliotheque par defaut" et plusieurs
# bibliotheques partagent BaseTemplate 101 (SiteAssets, Style Library...).
# On exclut donc les bibliotheques systeme puis on priorise celle dont
# RootFolder.Name vaut "Shared Documents" ou "Documents".
# Retourne l'objet liste (avec RootFolder charge).
function Get-DefaultDocumentLibrary {
    param(
        [Parameter(Mandatory)] $Connection
    )

    $docLibs = Get-PnPList -Connection $Connection -ErrorAction Stop |
               Where-Object {
                   $_.BaseTemplate -eq 101 -and
                   -not $_.Hidden -and
                   ($ExcludedLibInternalNames -notcontains $_.RootFolder.Name)
               }

    if (-not $docLibs) {
        throw "Aucune bibliotheque de documents (BaseTemplate 101) exploitable trouvee sur le site."
    }

    # 1) Priorite : RootFolder.Name = "Shared Documents" ou "Documents"
    $defaultLib = $docLibs |
                  Where-Object { $_.RootFolder.Name -in @('Shared Documents','Documents') } |
                  Select-Object -First 1

    # 2) Fallback : premiere bibliotheque non-systeme restante
    if (-not $defaultLib) {
        $defaultLib = $docLibs | Select-Object -First 1
    }

    return $defaultLib
}

# --- 3. Resolution de la racine selon TargetType ---------------------------
# Retourne un objet decrivant la racine a nettoyer :
#   SiteRelative   : chemin SITE-relative   (pour Get-PnPFolderItem)   ex: "Shared Documents/X"
#   ServerRelative : chemin SERVER-relative (pour Remove-PnPFolder)    ex: "/sites/web/Shared Documents/X"
#   LibSiteRel     : racine de lib SITE-relative                       ex: "Shared Documents"
#   WebServerRel   : ServerRelativeUrl du web                          ex: "/sites/web"
#   Leaf           : dernier segment                                   ex: "X"
function Resolve-TargetRoot {
    param(
        [Parameter(Mandatory)] [string]$TargetType,
        [Parameter(Mandatory)] [string]$TargetFolder,
        [Parameter(Mandatory)] $Connection
    )

    # Normalise les separateurs (\ -> /), supprime les slashes superflus
    $folder = ($TargetFolder -replace '\\','/' -replace '/+','/').Trim('/')

    $defaultLib  = Get-DefaultDocumentLibrary -Connection $Connection
    $libServer   = $defaultLib.RootFolder.ServerRelativeUrl.TrimEnd('/')   # ex: /sites/web/Shared Documents
    $libName     = $defaultLib.RootFolder.Name                              # ex: "Shared Documents"

    # ServerRelativeUrl du web (ex: /sites/web), obtenu en retirant le nom de lib
    # de la fin du chemin server-relative de la lib.
    $web         = Get-PnPWeb -Connection $Connection -ErrorAction Stop
    $webServer   = $web.ServerRelativeUrl.TrimEnd('/')                      # ex: /sites/web  (ou "" pour racine)

    switch -Wildcard ($TargetType.Trim()) {

        'SharePoint' {
            # Le TargetFolder contient DEJA "Shared Documents/..." ou "Documents/..."
            # -> on retire le prefixe lib s'il est present pour eviter le doublon.
            $clean = $folder -replace '^(Shared Documents|Documents)/', ''
        }

        'Teams-Channel*' {
            # Canal standard : dossier <Channel>/... DANS la lib du site d'equipe.
            $clean = $folder -replace '^(Shared Documents|Documents)/', ''
        }

        'Teams-Private-Channel' {
            # Canal prive : site collection dediee (URL deja specifique).
            $clean = $folder -replace '^(Shared Documents|Documents)/', ''
        }

        default {
            throw "TargetType inconnu : '$TargetType'"
        }
    }

    # Construction des deux representations de chemin
    $siteRel   = ("$libName/$clean"   -replace '/+','/').Trim('/')          # SITE-relative
    $serverRel = ("$libServer/$clean" -replace '/+','/').TrimEnd('/')       # SERVER-relative
    $libSiteRel = $libName

    [pscustomobject]@{
        SiteRelative   = $siteRel
        ServerRelative = $serverRel
        LibSiteRel     = $libSiteRel
        LibServerRel   = $libServer
        WebServerRel   = $webServer
        Leaf           = ($siteRel -split '/')[-1]
    }
}

# --- 4. Fonction recursive : suppression des sous-dossiers vides -----------
# Travaille en chemin SITE-relative pour Get-PnPFolderItem.
# Pour Remove-PnPFolder (qui exige du SERVER-relative), on prefixe WebServerRel.
# Un SEUL appel Get-PnPFolderItem par dossier (fichiers + dossiers), puis tri
# local. Decremente le compteur d'enfants au fil des suppressions plutot que de
# relire le serveur.
function Remove-EmptySubFolders {
    param(
        [Parameter(Mandatory)] $Connection,
        [Parameter(Mandatory)] [string]$FolderSiteRelativeUrl,
        [Parameter(Mandatory)] [string]$WebServerRelativeUrl,
        [Parameter(Mandatory)] [bool]$IsRoot,
        [Parameter(Mandatory)] [bool]$Simulate,
        [Parameter(Mandatory)] [string]$SiteUrl
    )

    $deletedCount = 0

    # Lister TOUS les items (un seul appel reseau) - chemin SITE-relative
    $items = @()
    try {
        $items = Get-PnPFolderItem -FolderSiteRelativeUrl $FolderSiteRelativeUrl `
                                   -Connection $Connection `
                                   -ErrorAction Stop |
                 Where-Object { $_.Name -and $_.Name -ne 'Forms' }
    } catch {
        Write-Log @log_params -Msg "[WARN] $SiteUrl  ->  $FolderSiteRelativeUrl : enumeration impossible ($($_.Exception.Message))" -Type WARN
        return 0
    }

    # Separer fichiers vs dossiers localement
    $folders = @($items | Where-Object { $_.GetType().Name -eq 'Folder' })
    $files   = @($items | Where-Object { $_.GetType().Name -eq 'File' })

    $childFolderCount = $folders.Count

    foreach ($child in $folders) {
        $childSiteRel = "$FolderSiteRelativeUrl/$($child.Name)"
        $childDeleted = Remove-EmptySubFolders -Connection $Connection `
                                               -FolderSiteRelativeUrl $childSiteRel `
                                               -WebServerRelativeUrl $WebServerRelativeUrl `
                                               -IsRoot $false `
                                               -Simulate $Simulate `
                                               -SiteUrl $SiteUrl
        $deletedCount += $childDeleted

        # Si l'enfant direct a ete supprime, on le retire du compte local.
        if ($childDeleted -gt 0 -and (Test-ChildDeleted -Connection $Connection -ChildSiteRel $childSiteRel)) {
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
            Write-Log @log_params -Msg "[WHATIF] $SiteUrl  ->  $FolderSiteRelativeUrl serait supprime" -Type INFO
            $deletedCount++
        } else {
            $leaf            = ($FolderSiteRelativeUrl -split '/')[-1]
            $parentSiteRel   = $FolderSiteRelativeUrl.Substring(0, $FolderSiteRelativeUrl.LastIndexOf('/'))
            # Remove-PnPFolder -Folder attend du SERVER-relative
            $parentServerRel = ("$WebServerRelativeUrl/$parentSiteRel" -replace '/+','/').TrimEnd('/')
            try {
                Remove-PnPFolder -Name $leaf `
                                 -Folder $parentServerRel `
                                 -Connection $Connection `
                                 -Force `
                                 -ErrorAction Stop
                Write-Log @log_params -Msg "[DEL] $SiteUrl  ->  $FolderSiteRelativeUrl supprime" -Type INFO
                $deletedCount++
            } catch {
                Write-Log @log_params -Msg "[KO] $SiteUrl  ->  $FolderSiteRelativeUrl : $($_.Exception.Message)" -Type ERROR
            }
        }
    }

    return $deletedCount
}

# Verifie qu'un dossier n'existe plus (apres tentative de suppression).
# Get-PnPFolder -Url accepte un chemin SITE-relative.
function Test-ChildDeleted {
    param(
        [Parameter(Mandatory)] $Connection,
        [Parameter(Mandatory)] [string]$ChildSiteRel
    )
    $exists = Get-PnPFolder -Url $ChildSiteRel -Connection $Connection -ErrorAction SilentlyContinue
    return (-not $exists)
}

# --- 5. Boucle principale --------------------------------------------------
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

        Write-Log @log_params -Msg "[LIB] $($t.Url)  ->  bibliotheque par defaut : $($root.LibServerRel)" -Type INFO

        # --- Garde-fous de securite ---
        if ($root.SiteRelative -eq $root.LibSiteRel) {
            Write-Log @log_params -Msg "[SKIP-SECURITE] $($t.Url)  ->  racine de bibliotheque protegee ($($root.SiteRelative))" -Type WARN
            [pscustomobject]@{ Type=$t.Type; Url=$t.Url; Folder=$t.Folder; Root=$root.SiteRelative; Deleted=0; Status='Skipped'; Message='Racine de bibliotheque protegee' }
            continue
        }
        if ($ProtectedLeafNames -contains $root.Leaf) {
            Write-Log @log_params -Msg "[SKIP-SECURITE] $($t.Url)  ->  dossier protege '$($root.Leaf)' ($($root.SiteRelative))" -Type WARN
            [pscustomobject]@{ Type=$t.Type; Url=$t.Url; Folder=$t.Folder; Root=$root.SiteRelative; Deleted=0; Status='Skipped'; Message="Dossier protege '$($root.Leaf)'" }
            continue
        }

        # Verifier l'existence de la racine (Get-PnPFolder -Url accepte du site-relative)
        $rootExists = Get-PnPFolder -Url $root.SiteRelative -Connection $conn -ErrorAction SilentlyContinue
        if (-not $rootExists) {
            Write-Log @log_params -Msg "[SKIP] $($t.Url)  ->  $($root.SiteRelative) inexistant" -Type WARN
            [pscustomobject]@{ Type=$t.Type; Url=$t.Url; Folder=$t.Folder; Root=$root.SiteRelative; Deleted=0; Status='Skipped'; Message='Racine inexistante' }
            continue
        }

        Write-Log @log_params -Msg "[..] $($t.Url)  ->  parcours de $($root.SiteRelative)" -Type INFO

        $deleted = Remove-EmptySubFolders -Connection $conn `
                                          -FolderSiteRelativeUrl $root.SiteRelative `
                                          -WebServerRelativeUrl $root.WebServerRel `
                                          -IsRoot $true `
                                          -Simulate $simulate `
                                          -SiteUrl $t.Url

        Write-Log @log_params -Msg "[OK] $($t.Url)  ->  $($root.SiteRelative) : $deleted dossier(s) vide(s)" -Type INFO

        [pscustomobject]@{
            Type    = $t.Type
            Url     = $t.Url
            Folder  = $t.Folder
            Root    = $root.SiteRelative
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

# --- 6. Export rapport + bilan ---------------------------------------------
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
