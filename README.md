# FileShare Assessment - phase 1

Ce dépôt contient une première phase de standardisation des scripts du dossier `Scripts Assessments` afin de les rendre plus réutilisables pour une squad Modern Workplace.

## Ce qui est prêt

- un module PowerShell `FileShareAssessment` importable sans lancer de scan ;
- un point d'entrée `Scripts Assessments/Start-Assessment.ps1` basé sur une configuration explicite ;
- une configuration d'exemple anonymisée ;
- une base de tests Pester et une configuration PSScriptAnalyzer ;
- une documentation en français pour l'utilisation, l'exploitation et l'audit technique.

## Prérequis

- Windows PowerShell 5.1 recommandé pour les scans sur serveurs Windows ;
- PowerShell 7 supporté pour le packaging et certains scripts, avec des différences documentées dans `docs/ARCHITECTURE.md` ;
- droits de lecture NTFS et partage sur les file shares à analyser ;
- espace disque local pour les exports (`Output/`) ;
- accès réseau vers les partages sources ;
- politique d'exécution permettant l'exécution de scripts signés ou internes selon votre standard d'équipe.

## Installation / récupération

```powershell
Set-Location 'Scripts Assessments'
Import-Module .\FileShareAssessment\FileShareAssessment.psd1 -Force
```

## Quick start

1. Copier ou adapter `Scripts Assessments/Config/FileShareAssessment.json`.
2. Adapter `Scripts Assessments/Config/FileShareMapping.example.csv` avec vos chemins et métadonnées.
3. Lancer un test de configuration :

```powershell
Import-Module .\Scripts Assessments\FileShareAssessment\FileShareAssessment.psd1 -Force
Test-FileShareAssessmentConfiguration -ConfigurationPath '.\Scripts Assessments\Config\FileShareAssessment.json'
```

4. Lancer le préflight sur la machine qui exécutera le scan. Il contrôle Windows/PowerShell, les UNC du mapping, l'espace disque et l'écriture dans la sortie ; un avertissement demande de vérifier les ACL de sortie :

```powershell
.\Scripts Assessments\Start-Assessment.ps1 -ConfigurationPath '.\Scripts Assessments\Config\FileShareAssessment.json' -PreflightOnly
```

5. Lancer l'assessment :

```powershell
.\Scripts Assessments\Start-Assessment.ps1 -ConfigurationPath '.\Scripts Assessments\Config\FileShareAssessment.json'
```

## Workflow d'assessment

Le point d'entrée standardisé exécute, dans l'ordre :

1. inventaire et dernier accès ;
2. permissions NTFS ;
3. extensions à auditer ;
4. doublons ;
5. chemins trop longs ;
6. rapport HTML final.

Les scripts historiques restent présents et peuvent toujours être appelés directement. La standardisation phase 1 les encapsule sans réécriture lourde.

## Entrées / sorties

- Entrées : JSON de configuration, CSV de mapping, credentials SMB optionnels.
- Sorties : structure `Output/<scope>/<horodatage>/` avec sous-dossiers `csv`, `metadata`, `logs`, `errors` et rapports HTML `_Reports`.

## Documentation

- `docs/TECHNICAL-AUDIT.md`
- `docs/ARCHITECTURE.md`
- `docs/OPERATIONS.md`
- `docs/TROUBLESHOOTING.md`

## Sécurité et confidentialité

- ne commitez pas les exports de scan ;
- utilisez les fichiers d'exemple fournis pour versionner la structure sans données client ;
- supprimez ou archivez les résultats de manière contrôlée après le projet ;
- limitez l'accès aux rapports aux personnes autorisées.
- utilisez un répertoire de sortie local ou partagé dont l'écriture est réservée à l'équipe projet ;
- les liens symboliques, jonctions et autres reparse points sont exclus du scan et consignés dans les exports d'erreurs.

## Contribution

- conservez la compatibilité PowerShell 5.1 lorsque c'est raisonnable ;
- privilégiez des changements incrémentaux et testables ;
- documentez toute dette technique identifiée avant une réécriture métier.
