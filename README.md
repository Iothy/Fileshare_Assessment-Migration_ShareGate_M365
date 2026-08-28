# FileShare Assessment - phase 1

Ce dépôt contient une standardisation des scripts du dossier `Scripts Assessments` pour préparer un assessment de file shares avant migration Microsoft 365.

## Ce qui est prêt

- un module PowerShell `FileShareAssessment` avec import/validation du mapping simplifié ;
- un point d'entrée `Scripts Assessments/Start-Assessment.ps1` basé sur une configuration explicite ;
- des exemples anonymisés ;
- des tests Pester et une configuration PSScriptAnalyzer ;
- une documentation opérateur en français.

## Quick start

1. Copier `Scripts Assessments/Config/FileShareMapping.example.csv`.
2. Respecter exactement ce modèle CSV UTF-8 `;` :

```text
SourcePath;TargetType;TargetSPOURL;TargetFolder;DateFilter (YYYY-DD-MM);Permissions
```

3. Signification des colonnes :
   - `SourcePath` : chemin UNC source obligatoire ;
   - `TargetType` : `SharePoint` ou `OneDrive` (casse libre en entrée) ;
   - `TargetSPOURL` : URL HTTPS cible, slash final supprimé ;
   - `TargetFolder` : dossier cible, vide autorisé, normalisé en `/` ;
   - `DateFilter (YYYY-DD-MM)` : optionnel, format strict `YYYY-DD-MM` (`2020-31-12` = 31 décembre 2020) ;
   - `Permissions` : `YES`/`NO` avec alias `Oui`/`Non`/`True`/`False`.

4. Exemple :

```text
SourcePath;TargetType;TargetSPOURL;TargetFolder;DateFilter (YYYY-DD-MM);Permissions
\\fs01.contoso.local\RH;SharePoint;https://contoso.sharepoint.com/sites/RH;Documents;2020-31-12;YES
\\fs01.contoso.local\RH\Paie;SharePoint;https://contoso.sharepoint.com/sites/RH;Documents/Paie;;Oui
\\fs01.contoso.local\Compta\Paie;SharePoint;https://contoso.sharepoint.com/sites/Finance;Archives/Paie;;NO
```

Cet exemple montre un chevauchement volontaire (`\\fs01.contoso.local\RH` et `\\fs01.contoso.local\RH\Paie`) et deux feuilles `Paie` sous des branches différentes : les identifiants restent distincts car ils sont dérivés du chemin complet.

## Sorties

Un run standard produit désormais :

```text
Output/<scope>/<yyyyMMdd_HHmmss>/
  Index_Assessment.html
  Rapport_Global.html
  manifest.json
  execution.log
  <SourceIdentifier>/
    Rapport_<SourceIdentifier>.html
    Synthese_<SourceIdentifier>.csv
    Inventaire_<SourceIdentifier>.csv
    InventaireDetail_<SourceIdentifier>.csv
    Permissions_<SourceIdentifier>.csv
    CheminsLongs_<SourceIdentifier>.csv
    Extensions_<SourceIdentifier>.csv
    Doublons_<SourceIdentifier>.csv
    AccesRefuses_<SourceIdentifier>.csv
    Execution_<SourceIdentifier>.log
```

`SourceIdentifier` est dérivé du chemin source complet (serveur, partage, sous-dossiers), avec caractères non sûrs neutralisés ; un suffixe de hash court n'est ajouté qu'en cas de collision.

## Validation

`Test-FileShareMapping` signale notamment :
- erreurs bloquantes d'en-têtes, doublons, chemins UNC invalides, URL non HTTPS, date invalide ou permissions invalides ;
- warnings de recouvrement de périmètre, destinations identiques, longueurs cible élevées et préflight UNC inaccessible.
