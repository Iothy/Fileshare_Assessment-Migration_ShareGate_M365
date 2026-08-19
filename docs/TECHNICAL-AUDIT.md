# Audit technique - `Scripts Assessments`

## Priorisation

### Bloquant

1. `Scripts Assessments/Start-Assessment.ps1` contenait uniquement des appels relatifs et une référence à `Get-InvalidCharacters.ps1`, absent du dossier courant.
2. Les fichiers de configuration versionnés contenaient des données client et des chemins nominaux.

### Important

1. Les scripts historiques dépendent de noms de modules et de messages encore marqués par un contexte client (`PrimaGAZ`).
2. Plusieurs scripts répètent des helpers identiques (`Get-SafeFileName`, `Get-NomFileShareFromPath`, `Get-ExecutionAccount`) au lieu d'un module commun.
3. Les sorties étaient partiellement structurées via `Modules/PrimaGAZ.Output.psm1`, mais aucun point d'entrée standardisé n'exploitait ce mécanisme.
4. `Export-AssessmentReport.ps1` est très volumineux et fortement couplé aux noms de fichiers CSV historiques.
5. Les scripts utilisent beaucoup `Write-Host` et des sorties ad hoc, ce qui limite l'automatisation.
6. Les dépendances SMB (`Get-SmbConnection`, `New-SmbMapping`, `Remove-SmbMapping`) impliquent une exécution Windows pour certains scénarios.

### Amélioration

1. Ajouter à terme un vrai découpage en fonctions privées/publiques pour les scripts métier volumineux.
2. Introduire des objets de sortie plus structurés et des codes d'état homogènes.
3. Mieux isoler les heuristiques spécifiques au reporting HTML dans des composants dédiés.
4. Étendre les tests unitaires aux fonctions métier communes après extraction.

## Compatibilité PowerShell

- **PowerShell 5.1** : cible prioritaire pour l'exécution sur serveurs Windows et pour les cmdlets SMB.
- **PowerShell 7** : le packaging, les tests et une partie des scripts fonctionnent, mais les scans SMB restent dépendants d'un hôte Windows et de ses modules.
- `Get-FileSharePermissions.ps1` mentionne explicitement des optimisations PS7+, avec fallback prévu ; cette coexistence doit être conservée tant que les algorithmes ne sont pas refactorés.

## Risques de performance

- scan récursif volumineux sur gros partages ;
- hashing des doublons potentiellement coûteux ;
- rapport HTML monolithique sur gros volumes ;
- écriture CSV multiple et tri en mémoire selon les contrôles.

## Recommandations phase suivante

1. Renommer progressivement les modules hérités pour supprimer le branding historique.
2. Extraire les helpers dupliqués dans le module `FileShareAssessment`.
3. Découper `Export-AssessmentReport.ps1` en fonctions thématiques testables.
4. Ajouter des tests unitaires ciblés sur les fonctions extraites et des jeux de données synthétiques.
