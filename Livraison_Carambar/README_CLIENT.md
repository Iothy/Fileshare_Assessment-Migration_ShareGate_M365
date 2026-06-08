# Assessment FileShare — Scripts d'analyse pré-migration M365 — Client Carambar

## Objectif

Ce package permet de réaliser un assessment complet d'un environnement FileShare avant migration vers Microsoft 365 (SharePoint Online / OneDrive), afin de préparer le cadrage, la stratégie de migration et le nettoyage des données.

## Prérequis

- Windows Server ou poste de travail avec accès réseau au FileShare
- **PowerShell 7** (obligatoire)
- Compte avec les permissions suivantes sur l'ensemble des partages à analyser :
  - **Read** (lecture des fichiers et dossiers)
  - **Read Permissions** (lecture des ACL/permissions NTFS)

  > ⚠️ Sans ces deux permissions, le script `Get-FileSharePermissions.ps1` ne fonctionnera pas.

- Optionnel : compte de service dédié (exécution via :
  `runas /netonly /user:<DOMAINE_AD>\<COMPTE_SERVICE_LECTURE> powershell.exe`)

## Configuration initiale

Renseigner le fichier `Scripts Assessments/Config/FileShareMapping.csv`.

- Colonnes obligatoires : `CheminUNC`, `NomFileShare`
- Colonnes optionnelles : `TypeUsage`, `CibleM365`, `Owner`, `EmailOwner`, `Description`

Bonnes pratiques :
- Lister **tous** les partages / chemins UNC à analyser
- Vérifier l'accessibilité réseau de chaque chemin UNC avant exécution

## Exécution

### Option 1 — Orchestration complète

Exécuter le script principal (lancement séquentiel des analyses incluses) :

```powershell
.\Start-Assessment.ps1
```

### Option 2 — Exécution script par script

```powershell
.\Get-FileShareInventory-WithLastAccess.ps1 -MappingCsv ".\Config\FileShareMapping.csv" -IncludeFileDetail
.\Get-FileSharePermissions.ps1 -MappingCsv ".\Config\FileShareMapping.csv" -ThrottleLimit 1
.\Get-BlockedExtensions.ps1 -MappingCsv ".\Config\FileShareMapping.csv"
.\Get-DuplicateFiles.ps1 -MappingCsv ".\Config\FileShareMapping.csv" -MinSizeMB 50
.\Get-InvalidCharacters.ps1 -MappingCsv ".\Config\FileShareMapping.csv"
.\Export-AssessmentReport.ps1 -CheminOutput ".\Output" -FileShareMapping ".\Config\FileShareMapping.csv"
```

## Description des scripts

### `Get-FileShareInventory-WithLastAccess.ps1`
Inventaire complet par partage : nombre de fichiers, dossiers, volumétrie et date de dernier accès.

Paramètre clé :
- `-IncludeFileDetail` : active un export détaillé fichier par fichier.

### `Get-FileSharePermissions.ps1`
Export des permissions NTFS pour analyse de la transposition vers M365.

Paramètres clés :
- `-ThrottleLimit` : niveau de parallélisme (défaut 8 en PowerShell 7)
- `-Depth` : profondeur d'analyse (`0` = illimité)

### `Get-BlockedExtensions.ps1`
Détecte les extensions historiquement bloquées par SharePoint Server.

Note : SharePoint Online ne bloque plus d'extensions par défaut ; ce script reste un audit préventif.

### `Get-DuplicateFiles.ps1`
Détecte les fichiers dupliqués via hash SHA256.

Paramètre clé :
- `-MinSizeMB` : taille minimale de calcul (défaut 100 MB, recommandé 50 MB pour 10 TO)

### `Get-InvalidCharacters.ps1`
Détecte les fichiers et dossiers dont le nom est incompatible avec SharePoint Online / OneDrive.

Contrôles réalisés :
- Caractères invalides : `" * : < > ? / \ |`
- Nom commençant par `#` ou `%`
- Nom se terminant par un espace ou un point
- Noms réservés Windows / SharePoint (CON, PRN, AUX, NUL, _vti_, forms, etc.)
- Fichiers système (`desktop.ini`), de verrouillage (`.lock`) et temporaires Office (`~$`)
- Nom dépassant 255 caractères

Paramètres clés :
- `-MappingCsv` : chemin du fichier FileShareMapping.csv (mode multi-chemins, recommandé)
- `-CheminUNC` : chemin unique à analyser (mode mono-chemin)
- `-OutputPath` : dossier de sortie des résultats (par défaut : `.\Output`)

#### Traitement par ShareGate lors de la migration

> ℹ️ Ce script est un **audit préventif**. Lors de la migration effective avec ShareGate, les caractères invalides sont traités automatiquement :
>
> - **Remplacement automatique** : ShareGate remplace par défaut les caractères invalides par un underscore `_` lors de la migration. Aucune action manuelle n'est requise pour ces fichiers.
> - **Personnalisation** : le caractère de remplacement peut être modifié dans ShareGate via *Settings → Migration → Special and illegal characters*.
> - **Cas particuliers `#` et `%`** : ces caractères sont désormais supportés par SharePoint Online, mais nécessitent une activation au niveau du tenant :
>   ```powershell
>   Set-SPOTenant -SpecialCharactersStateInFileFolderNames Allowed
>   ```
>   Sans cette activation, ShareGate les remplacera automatiquement.
> - **Fichiers système** (`desktop.ini`, `.lock`, `~$*`) : ces fichiers sont **skippés nativement** par ShareGate et ne seront pas migrés (aucune action requise).
>
> **Sévérités du script** :
> - `ERROR` : l'item ne migrera pas sans correction ou remplacement automatique (caractères invalides, noms réservés, noms > 255 car.)
> - `WARN` : l'item migrera mais avec un risque selon la configuration du tenant (préfixes `#` / `%`)
> - `INFO` : fichier skippé nativement par ShareGate (aucune action requise)

### `Export-AssessmentReport.ps1`
Génère un rapport HTML décisionnel consolidé à partir des CSV produits.

#### Prérequis spécifiques

- Le module `Modules\Carambar.Assessment.psm1` doit être présent dans le dossier `Scripts Assessments/Modules/`
- Le dossier `Output/` doit exister et contenir les fichiers CSV générés par les scripts d'assessment précédents (inventaire, permissions, extensions bloquées, doublons)
- Le fichier `Config\FileShareMapping.csv` doit être renseigné (pour le mode enrichi par FileShare)

> ⚠️ Ce script doit être exécuté **après** les autres scripts d'analyse. Sans les CSV dans le dossier `Output/`, le rapport ne contiendra aucune donnée.

Paramètres clés :
- `-CheminOutput` : chemin du dossier Output contenant les CSV (par défaut : `.\Output`)
- `-FileShareMapping` : chemin du fichier FileShareMapping.csv (active le mode multi-rapports par FileShare)
- `-ReportOutputPath` : dossier de destination des rapports HTML (par défaut : `Output/_Reports/`)
- `-SplitByLevel1` : génère des rapports séparés par dossier de niveau 1

## Sorties

- Création automatique du dossier `Output/`
- Fichiers CSV par analyse
- Fichiers de métadonnées JSON
- Logs d'exécution
- Rapport HTML consolidé (via `Export-AssessmentReport.ps1`)

## Analyse des chemins trop longs

Cette analyse n'est **pas incluse** dans ce package.

Raison : le calcul pertinent nécessite de connaître les URLs SharePoint de destination (préfixe URL cible).

Action attendue : fournir la liste des sites SharePoint cibles avec leurs URLs pour permettre la réalisation de cette analyse dans un second temps.

## Support

En cas de problème, contacter l'équipe de migration.
