# Dépannage

## Erreurs fréquentes

### Le module ne s'importe pas

- vérifier le chemin du dépôt ;
- vérifier la politique d'exécution ;
- vérifier que les fichiers `FileShareAssessment.psd1` et `.psm1` sont présents.

### La configuration est invalide

Exécuter :

```powershell
Test-FileShareAssessmentConfiguration -ConfigurationPath '.\Scripts Assessments\Config\FileShareAssessment.json'
```

Puis corriger le chemin du CSV de mapping ou d'un script historique référencé.

### Les cmdlets SMB sont absentes

Les fonctions de credential SMB historiques nécessitent un environnement Windows avec les cmdlets SMB adéquates. Sous PowerShell 7 Linux/macOS, le packaging se teste mais le scan réel doit être exécuté sur Windows.

### Les performances sont insuffisantes

- réduire le périmètre du mapping CSV ;
- augmenter progressivement les seuils de doublons ;
- éviter `IncludeFileDetail` sur un périmètre trop large lors des premiers runs ;
- exécuter les scans au plus près du file server.

## Limitations connues

- le reporting HTML reste fortement couplé aux noms de fichiers historiques ;
- les modules hérités portent encore un nom client ;
- l'ancienne référence à `Get-InvalidCharacters.ps1` n'est pas réintroduite dans le nouveau point d'entrée car le script n'est pas présent dans ce clone.
