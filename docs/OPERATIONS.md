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
- `TargetType` accepte `SharePoint` / `OneDrive` ;
- `Permissions` accepte `YES` / `NO` et alias `Oui` / `Non` / `True` / `False` ;
- `DateFilter (YYYY-DD-MM)` utilise volontairement `YYYY-DD-MM` (`2020-31-12` = 31 décembre 2020).

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
- erreur : `TargetSPOURL` non HTTPS ;
- erreur : `2020-12-31` rejeté car le format attendu est `YYYY-DD-MM` ;
- warning : recouvrement entre `\\fs01.contoso.local\RH` et `\\fs01.contoso.local\RH\Paie` ;
- warning : plusieurs sources vers la même destination normalisée.

## Bonnes pratiques sécurité / RGPD

- minimiser les métadonnées collectées au strict besoin du projet ;
- stocker les exports sur un emplacement restreint ;
- définir une durée de rétention ;
- supprimer les exports en fin de mission ;
- éviter de diffuser les rapports contenant noms, ACL ou arborescences à un public non autorisé.
