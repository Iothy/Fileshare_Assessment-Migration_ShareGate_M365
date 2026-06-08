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

## Guide de lecture des résultats — Permissions NTFS (`Get-FileSharePermissions.ps1`)

### Fichiers générés

Le script produit les fichiers suivants dans le dossier `Output/Permissions/` :

| Fichier | Description |
|---------|-------------|
| `Permissions_NTFS_<NomFS>_<timestamp>.csv` | Export complet des ACL par FileShare |
| `Permissions_NTFS_<timestamp>.csv` | CSV consolidé (tous les FileShares) |
| `Permissions_InheritanceBroken_<timestamp>.csv` | Liste des dossiers avec héritage cassé |
| `AccessDenied_<NomFS>_<timestamp>.csv` | Dossiers refusés lors de l'audit (droits insuffisants) |

### Structure du CSV principal — Colonnes expliquées

| Colonne | Signification | Comment lire |
|---------|---------------|--------------|
| `CheminDossier` | Chemin UNC complet du dossier analysé | Permet d'identifier précisément le dossier concerné |
| `Proprietaire` | Propriétaire NTFS du dossier | Souvent le créateur original — utile pour identifier les responsables |
| `Identite` | Utilisateur ou groupe AD qui possède l'accès | **Colonne clé** : identifie QUI a accès au dossier |
| `TypeAcces` | `Allow` (autoriser) ou `Deny` (refuser) | Les `Deny` sont prioritaires et bloquent l'accès même si un `Allow` existe |
| `Droits` | Niveau de permission accordé | Voir tableau des droits ci-dessous |
| `EstHerite` | `True` = hérité du dossier parent, `False` = défini explicitement | Les ACE explicites (`False`) sont les plus importantes à analyser |
| `InheritanceFlags` | Comment la permission se propage aux enfants | `ContainerInherit` = s'applique aux sous-dossiers, `ObjectInherit` = s'applique aux fichiers |
| `PropagationFlags` | Contrôle de propagation avancé | `None` = s'applique ici et aux enfants, `InheritOnly` = ne s'applique qu'aux enfants |
| `NomFileShare` | Nom du FileShare source | Permet de filtrer les résultats par partage |

### Comprendre les droits NTFS

| Droit | Signification | Équivalent simplifié |
|-------|---------------|---------------------|
| `FullControl` | Contrôle total (lecture, écriture, modification, suppression, modification des permissions) | Administrateur du dossier |
| `Modify` | Lecture + écriture + suppression (sans modifier les permissions) | Contributeur avancé |
| `ReadAndExecute, Synchronize` | Lecture seule + exécution | Lecteur |
| `Write` | Écriture / ajout de contenu | Contributeur limité |
| `Read` | Lecture seule | Lecteur strict |
| `ListDirectory` | Lister le contenu d'un dossier uniquement | Navigation |

### Comprendre l'héritage NTFS

L'héritage est un concept fondamental pour la lecture des résultats :

- **Héritage activé** (par défaut) : les permissions du dossier parent se propagent automatiquement aux sous-dossiers et fichiers. C'est le comportement normal.
- **Héritage cassé** (`AreAccessRulesProtected = True`) : le dossier a ses propres permissions, indépendantes du parent. C'est une **rupture d'héritage**.

#### Pourquoi les ruptures d'héritage sont importantes pour la migration ?

Dans SharePoint Online / OneDrive, chaque rupture d'héritage :
- **Crée un périmètre de sécurité unique** qui doit être reproduit côté M365
- **Peut nécessiter un site ou une bibliothèque dédiée** si les permissions sont très différentes du parent
- **Complexifie la gouvernance** post-migration (plus de ruptures = plus de maintenance)

#### Décisions à prendre par le client

| Situation | Action recommandée |
|-----------|-------------------|
| Dossier avec héritage cassé + permissions très restreintes | Envisager un **site SharePoint dédié** ou une **bibliothèque avec permissions uniques** |
| Dossier avec héritage cassé + même groupe que le parent | **Rétablir l'héritage** avant migration (nettoyage) |
| Nombreuses ruptures dans une même arborescence | **Restructurer** l'arborescence pour simplifier le modèle de permissions |
| Groupes AD obsolètes ou inconnus dans les ACL | **Nettoyer les ACL** ou **mapper les groupes** vers Entra ID avant migration |

### Comprendre les identités (colonne `Identite`)

| Type d'identité | Exemple | Signification |
|-----------------|---------|---------------|
| Groupe AD domaine | `DOMAINE\GRP_Finance_RW` | Groupe Active Directory — doit être mappé vers un groupe Entra ID / M365 |
| Utilisateur nominatif | `DOMAINE\jean.dupont` | Permission individuelle — à valider (devrait idéalement passer par un groupe) |
| Groupes built-in | `BUILTIN\Administrators` | Groupe Windows local — ne sera **pas** migré tel quel |
| `NT AUTHORITY\SYSTEM` | Compte système | Ignoré lors de la migration (pas de sens dans M365) |
| `Everyone` / `Tout le monde` | Accès universel | ⚠️ **Point d'attention critique** : signifie que tout le monde a accès — à sécuriser dans M365 |
| SID non résolu | `S-1-5-21-...` | Compte supprimé ou domaine inaccessible — à nettoyer |

### Plan d'action basé sur les résultats

1. **Identifier les ruptures d'héritage** → Fichier `Permissions_InheritanceBroken_*.csv`
   - Arbitrer pour chaque rupture : conserver (site/biblio dédié) ou rétablir l'héritage

2. **Mapper les groupes AD** → Colonne `Identite` (filtrer `EstHerite = False`)
   - Lister tous les groupes AD uniques avec des ACE explicites
   - Définir leur équivalent dans Entra ID (groupe M365, groupe de sécurité, etc.)

3. **Nettoyer les SID orphelins** → Identités au format `S-1-5-21-...`
   - Supprimer les entrées ACL pour les comptes qui n'existent plus

4. **Traiter les dossiers refusés** → Fichier `AccessDenied_*.csv`
   - Décider si ces dossiers doivent être migrés (nécessite un accès avec des droits supérieurs)

5. **Simplifier les permissions individuelles** → Identités avec `\prenom.nom`
   - Remplacer par des groupes AD quand c'est possible avant migration

### Correspondance NTFS → SharePoint Online

| Permission NTFS | Équivalent SharePoint Online |
|-----------------|------------------------------|
| `FullControl` | Contrôle total (Propriétaire du site) |
| `Modify` | Contribuer (Membre du site) |
| `ReadAndExecute` | Lire (Visiteur du site) |
| `Write` sans `Read` | ⚠️ N'existe pas nativement dans SPO — sera ajusté |
| `Deny` explicite | ⚠️ Pas de concept `Deny` dans SPO — doit être géré par exclusion |

> ℹ️ **ShareGate et la migration des permissions** : ShareGate peut migrer les permissions NTFS vers SharePoint Online en mappant les groupes AD vers leurs équivalents Entra ID. Cependant, les ruptures d'héritage et les permissions complexes nécessitent une planification préalable pour garantir un résultat cohérent dans M365.

## Analyse des chemins trop longs

Cette analyse n'est **pas incluse** dans ce package.

Raison : le calcul pertinent nécessite de connaître les URLs SharePoint de destination (préfixe URL cible).

Action attendue : fournir la liste des sites SharePoint cibles avec leurs URLs pour permettre la réalisation de cette analyse dans un second temps.

## Support

En cas de problème, contacter l'équipe de migration.
