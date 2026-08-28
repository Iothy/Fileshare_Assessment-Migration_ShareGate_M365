# Guide d'exploitation

## Installation

```powershell
Import-Module .\Scripts Assessments\FileShareAssessment\FileShareAssessment.psd1 -Force
```

## Validation du mapping

```powershell
Test-FileShareMapping -Path '.\Scripts Assessments\Config\FileShareMapping.example.csv'
```

Le CSV doit être en UTF-8, séparé par `;`, et contenir exactement :

```text
SourcePath;TargetType;TargetSPOURL;TargetFolder;DateFilter (YYYY-DD-MM);Permissions
```

Rappels :
- seul `SourcePath` est obligatoire pour lancer l'assessment général ;
- `TargetType` accepte `SharePoint`, `OneDrive`, `Teams-Channel General` ou `Teams-Private-Channel` lorsqu'il est renseigné ;
- si `TargetSPOURL` est vide, le contrôle chemins longs est marqué `Skipped` pour cette source, sans bloquer les autres contrôles ;
- `Permissions` accepte `YES` / `NO` et alias `Oui` / `Non` / `True` / `False` ;
- `DateFilter (YYYY-DD-MM)` utilise volontairement `YYYY-DD-MM` (`2020-31-12` = 31 décembre 2020).

## Validation de configuration

```powershell
Test-FileShareAssessmentConfiguration -ConfigurationPath '.\Scripts Assessments\Config\FileShareAssessment.json'
```

## Préflight obligatoire

Exécuter depuis une machine Windows ayant la connectivité SMB vers toutes les sources :

```powershell
.\Scripts Assessments\Start-Assessment.ps1 -ConfigurationPath '.\Scripts Assessments\Config\FileShareAssessment.json' -PreflightOnly
```

Corriger tout contrôle `Fail` avant le lancement. Un contrôle `Warn` sur les ACL de sortie signifie que les exports pourraient être accessibles à un public trop large : sélectionner un emplacement à accès restreint.

## Lancement d'un assessment complet

```powershell
.\Scripts Assessments\Start-Assessment.ps1 -ConfigurationPath '.\Scripts Assessments\Config\FileShareAssessment.json'
```

## Sorties attendues

```text
Output/<scope>/<yyyyMMdd_HHmmss>/
  Index_Assessment.html
  Rapport_Global.html
  manifest.json
  execution.log
  <SourceIdentifier>/...
```

Chaque sous-dossier source contient les CSV, le rapport HTML et `Execution_<SourceIdentifier>.log`.

## Erreurs et warnings fréquents

- erreur : en-têtes manquants, renommés, supplémentaires ou réordonnés ;
- erreur : `SourcePath` vide, non UNC, sans serveur ou sans partage ;
- erreur : `TargetSPOURL` non HTTPS lorsqu'il est renseigné ;
- erreur : `2020-12-31` rejeté car le format attendu est `YYYY-DD-MM` ;
- warning : recouvrement entre `\\fs01.contoso.local\RH` et `\\fs01.contoso.local\RH\Paie` ;
- warning : plusieurs sources vers la même destination normalisée.

## Chemins longs

Le contrôle `CheminsLongs` est dépendant d'une cible M365. Si le client fournit seulement des chemins UNC dans le mapping, le contrôle est ignoré proprement (`Skipped`) pour ces lignes.

Avec une cible renseignée, les paramètres par défaut sont :
- `SpoPathLimit = 400` : limite SharePoint Online sur le chemin cible décodé complet, nom du fichier inclus ;
- `WindowsOfficePathLimit = 256` ;
- `EstimatedLocalPrefixLength = 96`, hypothèse projet configurable pour le préfixe local OneDrive ;
- budget relatif Windows/Office par défaut : `256 - 96 = 160`.

Références Microsoft utilisées : limites OneDrive/SharePoint (`support.microsoft.com/office/invalid-file-names-and-file-types-in-onedrive-and-sharepoint-64883a5d-228e-48f5-b3d2-eb39e07630fa`) et limitation Windows MAX_PATH (`learn.microsoft.com/windows/win32/fileio/maximum-file-path-limitation`). Le préfixe 96 caractères est une hypothèse projet, pas une limite Microsoft.

## Bonnes pratiques sécurité / RGPD

- minimiser les métadonnées collectées au strict besoin du projet ;
- stocker les exports sur un emplacement restreint ;
- définir une durée de rétention ;
- supprimer les exports en fin de mission ;
- éviter de diffuser les rapports contenant noms, ACL ou arborescences à un public non autorisé.
- les jonctions, liens symboliques et reparse points sont exclus et tracés afin d'éviter les sorties de périmètre et les boucles de scan ;
- le mécanisme de credential SMB ne ferme pas les connexions SMB existantes : fermer explicitement une session concurrente avant d'utiliser un autre compte.
