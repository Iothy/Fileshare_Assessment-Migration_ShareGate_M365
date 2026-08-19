# Guide d'exploitation

## Installation

```powershell
Import-Module .\Scripts Assessments\FileShareAssessment\FileShareAssessment.psd1 -Force
```

## Validation de configuration

```powershell
Test-FileShareAssessmentConfiguration -ConfigurationPath '.\Scripts Assessments\Config\FileShareAssessment.json'
```

## Lancement d'un assessment complet

```powershell
.\Scripts Assessments\Start-Assessment.ps1 -ConfigurationPath '.\Scripts Assessments\Config\FileShareAssessment.json'
```

## Contrôles couverts

- inventaire global et dernier accès ;
- permissions NTFS ;
- extensions à auditer ;
- doublons ;
- chemins trop longs ;
- rapport HTML final.

## Fichiers produits

- `csv/` : exports consolidés ;
- `metadata/` : résumés JSON ;
- `logs/` : journaux d'exécution ;
- `errors/` : exports d'erreurs d'accès ;
- `_Reports/` : rapports HTML et annexes.

## Bonnes pratiques sécurité / RGPD

- minimiser les métadonnées collectées au strict besoin du projet ;
- stocker les exports sur un emplacement restreint ;
- définir une durée de rétention ;
- supprimer les exports en fin de mission ;
- éviter de diffuser les rapports contenant noms, ACL ou arborescences à un public non autorisé.
