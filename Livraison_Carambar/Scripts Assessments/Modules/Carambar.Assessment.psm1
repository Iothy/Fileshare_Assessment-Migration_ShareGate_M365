<#
.SYNOPSIS
    Module commun pour les scripts d'assessment Carambar.

.DESCRIPTION
    Ce module centralise les fonctions utilitaires partagées par tous les scripts
    d'assessment de la Phase 01 :
    - Write-Log                 : Logging console + fichier avec niveaux colorés
    - Set-LogFile               : Définit le fichier de log actif (console + fichier)
    - Get-LogFile               : Retourne le chemin du fichier de log actif
    - Write-ExecutionLog        : Écriture du fichier .log d'exécution (résumé final)
    - Initialize-OutputPath     : Vérification et création du dossier de sortie
    - Test-SourcePath           : Vérification de l'accessibilité du chemin source
    - Invoke-SafeRecursiveScan  : Énumération récursive 100% résiliente (remplace Get-ChildItem -Recurse)

.NOTES
    Projet  : Carambar - Migration FileShare vers M365
    Phase   : 01 - Assessment
#>

# Variable module pour stocker le LogPath actif
$script:CurrentLogFile = $null

function Set-LogFile {
    <#
    .SYNOPSIS
        Définit le fichier de log actif. Tous les appels Write-Log suivants écriront dans ce fichier.
    .PARAMETER Path
        Chemin complet du fichier .log à créer/utiliser.
    #>
    param([string]$Path)
    $script:CurrentLogFile = $Path
    if ($Path -and -not (Test-Path $Path)) {
        # Créer le fichier avec en-tête UTF-8 BOM (lisible dans Notepad sans corruption d'accents)
        $utf8Bom = New-Object System.Text.UTF8Encoding($true)
        $header = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [INFO] === Démarrage logging vers $Path ==="
        [System.IO.File]::WriteAllText($Path, $header + [Environment]::NewLine, $utf8Bom)
    }
}

function Get-LogFile {
    <#
    .SYNOPSIS
        Retourne le chemin du fichier de log actif, ou $null si aucun n'est défini.
    #>
    return $script:CurrentLogFile
}

function Write-Log {
    <#
    .SYNOPSIS
        Écrit un message horodaté sur la console ET dans le fichier de log actif (si Set-LogFile a été appelé).
    .PARAMETER Message
        Texte du message à logger.
    .PARAMETER Level
        Niveau de log : INFO, WARN, ERROR, SUCCESS, DEBUG. Défaut : INFO.
    .PARAMETER NoConsole
        Si présent, n'écrit pas sur la console (fichier uniquement).
    #>
    param(
        [string]$Message,
        [string]$Level = "INFO",
        [switch]$NoConsole
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "[$timestamp] [$Level] $Message"

    # Écriture console (sauf si -NoConsole)
    if (-not $NoConsole) {
        switch ($Level) {
            "INFO"    { Write-Host $line -ForegroundColor Cyan }
            "WARN"    { Write-Host $line -ForegroundColor Yellow }
            "ERROR"   { Write-Host $line -ForegroundColor Red }
            "SUCCESS" { Write-Host "[$timestamp] [OK] $Message" -ForegroundColor Green }
            "DEBUG"   { Write-Host $line -ForegroundColor DarkGray }
            default   { Write-Host "[$timestamp] [INFO] $Message" }
        }
    }

    # Écriture fichier (toujours, si LogPath défini)
    if ($script:CurrentLogFile) {
        try {
            # Append thread-safe via [System.IO.File]::AppendAllText
            [System.IO.File]::AppendAllText($script:CurrentLogFile, $line + [Environment]::NewLine, [System.Text.UTF8Encoding]::new($true))
        } catch {
            # Ne jamais crasher à cause du logging
            Write-Host "[$timestamp] [LOG-ERROR] Impossible d'écrire dans $script:CurrentLogFile : $($_.Exception.Message)" -ForegroundColor DarkRed
        }
    }
}

function Write-ExecutionLog {
    param(
        [string]$LogPath,
        [string]$ScriptName,
        [datetime]$StartTime,
        [datetime]$EndTime,
        [string]$CheminUNC,
        [string]$OutputCsv,
        [int]$ElementsAnalyses,
        [int]$ResultatsTrouves,
        [int]$WarningCount,
        [int]$ErrorCount,
        [string]$Status
    )
    $duration = $EndTime - $StartTime
@"
ScriptName          : $ScriptName
StartTime           : $($StartTime.ToString('yyyy-MM-dd HH:mm:ss'))
EndTime             : $($EndTime.ToString('yyyy-MM-dd HH:mm:ss'))
Duration            : $duration
CheminUNC           : $CheminUNC
OutputCsv           : $OutputCsv
ElementsAnalyses    : $ElementsAnalyses
ResultatsTrouves    : $ResultatsTrouves
WarningCount        : $WarningCount
ErrorCount          : $ErrorCount
Status              : $Status
"@ | Out-File -FilePath $LogPath -Encoding UTF8
}

function Initialize-OutputPath {
    param([string]$Path)
    if (-not (Test-Path -Path $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
        Write-Log "Dossier de sortie créé : $Path"
    }
}

function Test-SourcePath {
    param([string]$Path)
    if (-not (Test-Path -Path $Path)) {
        throw "Le chemin '$Path' n'existe pas ou n'est pas accessible."
    }
}

function Invoke-SafeRecursiveScan {
    <#
    .SYNOPSIS
        Énumère récursivement les entrées d'un chemin (fichiers + dossiers) de manière 100% résiliente.

    .DESCRIPTION
        Implémentation stack-based via [System.IO.Directory]::EnumerateFileSystemEntries().
        Chaque erreur (énumération dossier, lecture propriété) est :
        1. Loggée dans le fichier .log via Write-Log [ERROR/WARN]
        2. Ajoutée à la collection $ErrorCollection passée en référence (pour reporting CSV)
        3. Le scan continue sur l'objet suivant — JAMAIS de crash

    .PARAMETER RootPath
        Chemin racine à énumérer.

    .PARAMETER FilesOnly
        Si présent, ne renvoie que les fichiers (pas les dossiers).

    .PARAMETER DirectoriesOnly
        Si présent, ne renvoie que les dossiers (pas les fichiers).

    .PARAMETER ErrorCollection
        Liste [System.Collections.Generic.List[object]] où sont accumulées les erreurs détaillées.

    .PARAMETER ProgressActivity
        (optionnel) Activité affichée dans Write-Progress tous les 1000 items.

    .PARAMETER NomFileShare
        Nom du FileShare en cours d'analyse (pour les messages de log et les rapports d'erreur).

    .OUTPUTS
        Stream de [System.IO.FileSystemInfo] (FileInfo ou DirectoryInfo).

    .EXAMPLE
        $errors = New-Object 'System.Collections.Generic.List[object]'
        Invoke-SafeRecursiveScan -RootPath "\\srv\share" -ErrorCollection $errors | ForEach-Object {
            # traiter $_
        }
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$RootPath,
        [switch]$FilesOnly,
        [switch]$DirectoriesOnly,
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.List[object]]$ErrorCollection,
        [string]$ProgressActivity = $null,
        [string]$NomFileShare = ""
    )

    $stack = New-Object 'System.Collections.Generic.Stack[string]'
    $stack.Push($RootPath)
    $itemCount = 0

    while ($stack.Count -gt 0) {
        $currentDir = $stack.Pop()

        # Énumération non-récursive du dossier courant (isolée en try/catch)
        $entries = $null
        try {
            $entries = [System.IO.Directory]::EnumerateFileSystemEntries($currentDir, '*', [System.IO.SearchOption]::TopDirectoryOnly)
        } catch [System.UnauthorizedAccessException] {
            $errEntry = [PSCustomObject]@{
                NomFileShare  = $NomFileShare
                Chemin        = $currentDir
                TypeErreur    = 'AccessDenied'
                ExceptionType = $_.Exception.GetType().Name
                MessageErreur = $_.Exception.Message
                DateDetection = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            }
            $ErrorCollection.Add($errEntry)
            Write-Log "ACCESS DENIED sur dossier : $currentDir - $($_.Exception.Message)" "WARN"
            continue
        } catch {
            $errEntry = [PSCustomObject]@{
                NomFileShare  = $NomFileShare
                Chemin        = $currentDir
                TypeErreur    = 'EnumerationError'
                ExceptionType = $_.Exception.GetType().Name
                MessageErreur = $_.Exception.Message
                DateDetection = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            }
            $ErrorCollection.Add($errEntry)
            Write-Log "Erreur énumération dossier : $currentDir [$($_.Exception.GetType().Name)] $($_.Exception.Message)" "ERROR"
            continue
        }

        foreach ($entry in $entries) {
            $itemCount++

            if ($ProgressActivity -and ($itemCount % 1000 -eq 0)) {
                Write-Progress -Activity $ProgressActivity -Status "[$NomFileShare] $itemCount items énumérés..."
            }

            # Récupération sécurisée des propriétés (FileInfo ou DirectoryInfo)
            $info = $null
            $isDirectory = $false
            try {
                # Détection rapide via attributs (1 appel système au lieu de 2)
                $attrs = [System.IO.File]::GetAttributes($entry)
                $isDirectory = ($attrs -band [System.IO.FileAttributes]::Directory) -eq [System.IO.FileAttributes]::Directory

                if ($isDirectory) {
                    $info = [System.IO.DirectoryInfo]::new($entry)
                } else {
                    $info = [System.IO.FileInfo]::new($entry)
                }
            } catch {
                $errEntry = [PSCustomObject]@{
                    NomFileShare  = $NomFileShare
                    Chemin        = $entry
                    TypeErreur    = if ($_.Exception.Message -match 'denied|refus|Unauthorized|non autoris') { 'AccessDenied' } else { 'AccessError' }
                    ExceptionType = $_.Exception.GetType().Name
                    MessageErreur = $_.Exception.Message
                    DateDetection = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                }
                $ErrorCollection.Add($errEntry)
                Write-Log "Impossible d'ouvrir l'objet : $entry [$($_.Exception.GetType().Name)] $($_.Exception.Message)" "WARN"
                continue
            }

            # Empiler les sous-dossiers pour traversée
            if ($isDirectory) {
                $stack.Push($entry)
            }

            # Filtrage selon -FilesOnly / -DirectoriesOnly
            if ($FilesOnly -and $isDirectory) { continue }
            if ($DirectoriesOnly -and -not $isDirectory) { continue }

            # Émettre l'objet dans le pipeline
            $info
        }
    }
}

Export-ModuleMember -Function Write-Log, Write-ExecutionLog, Initialize-OutputPath, Test-SourcePath, Set-LogFile, Get-LogFile, Invoke-SafeRecursiveScan
