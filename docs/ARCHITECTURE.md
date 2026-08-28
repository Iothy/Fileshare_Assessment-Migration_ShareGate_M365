# Architecture - phase 1

## Structure introduite

- `Scripts Assessments/FileShareAssessment/`
  - `FileShareAssessment.psd1`
  - `FileShareAssessment.psm1`
  - `Public/` : `Import-FileShareAssessmentConfiguration`, `Test-FileShareAssessmentConfiguration`, `Invoke-FileShareAssessment`, `Import-FileShareMapping`, `Test-FileShareMapping`, `Resolve-FileShareSourceMetadata` ;
  - `Private/` : helpers de normalisation UNC / URL / dossier / date / permissions et de dérivation des identifiants de source.
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

Le mapping d'entrée est désormais réduit à 6 colonnes métier. Les métadonnées historiques (`NomFileShare`, `TypeUsage`, `Owner`, etc.) ne sont plus stockées dans le CSV : elles sont dérivées du `SourcePath` au moment de l'import.

`SourceIdentifier` est calculé à partir du chemin UNC complet, rendu compatible système de fichiers et suffixé par un hash court uniquement en cas de collision.

## Sorties

Les runs utilisent une structure plate à la racine du run, avec un sous-dossier par `SourceIdentifier`. Les rapports HTML et CSV sont donc regroupés par source au lieu d'être éclatés entre `csv/`, `metadata/`, `logs/` et `errors/`.

## Compatibilité

- Le module `FileShareAssessment` est importable sous PowerShell 5.1 et 7.
- Les scripts historiques restent appelables directement (mode `-CheminUNC`/`-MappingCsv` selon le script).
- Les opérations SMB restent dépendantes d'un environnement Windows avec les cmdlets adéquates.

## Dette technique volontairement conservée

- modules hérités `PrimaGAZ.*` non renommés dans cette PR pour éviter une régression large ;
- duplication de helpers non supprimée dans les scripts historiques.
