# Architecture - phase 1

## Structure introduite

- `Scripts Assessments/FileShareAssessment/`
  - `FileShareAssessment.psd1`
  - `FileShareAssessment.psm1`
  - `Public/`
  - `Private/`
- `Scripts Assessments/Config/`
  - `FileShareAssessment.json`
  - `FileShareMapping.example.csv`
- `Scripts Assessments/Tests/`

## Principe

La phase 1 encapsule les scripts historiques existants au lieu de les réécrire :

- `Invoke-FileShareAssessment` charge une configuration explicite ;
- crée un run horodaté via `Modules/PrimaGAZ.Output.psm1` ;
- exécute les scripts activés dans l'ordre défini ;
- enregistre un manifest et un historique de run.

## Compatibilité

- Le module `FileShareAssessment` est importable sous PowerShell 5.1 et 7.
- Les scripts historiques restent appelables directement.
- Les opérations SMB restent dépendantes d'un environnement Windows avec les cmdlets adéquates.

## Dette technique volontairement conservée

- modules hérités `PrimaGAZ.*` non renommés dans cette PR pour éviter une régression large ;
- `Export-AssessmentReport.ps1` conservé tel quel hors encapsulation ;
- duplication de helpers non supprimée dans les scripts historiques.
