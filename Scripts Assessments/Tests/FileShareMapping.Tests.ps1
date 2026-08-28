Describe 'FileShare mapping simplifié' {
    BeforeAll {
        $script:repoRoot = Split-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -Parent
        $script:modulePath = Join-Path -Path $script:repoRoot -ChildPath 'Scripts Assessments/FileShareAssessment/FileShareAssessment.psd1'
        $script:outputModulePath = Join-Path -Path $script:repoRoot -ChildPath 'Scripts Assessments/Modules/PrimaGAZ.Output.psm1'
        $script:reportScriptPath = Join-Path -Path $script:repoRoot -ChildPath 'Scripts Assessments/Export-AssessmentReport.ps1'
        $script:startAssessmentPath = Join-Path -Path $script:repoRoot -ChildPath 'Scripts Assessments/Start-Assessment.ps1'
        Import-Module -Name $script:modulePath -Force
        Import-Module -Name $script:outputModulePath -Force

        function New-TestFolder {
            $folder = Join-Path ([System.IO.Path]::GetTempPath()) ([guid]::NewGuid().ToString())
            New-Item -ItemType Directory -Path $folder -Force | Out-Null
            return $folder
        }

        function New-MappingFile {
            param(
                [Parameter(Mandatory)]
                [string]$Folder,
                [Parameter(Mandatory)]
                [string[]]$Rows
            )

            $path = Join-Path $Folder 'mapping.csv'
            @(
                'SourcePath;TargetType;TargetSPOURL;TargetFolder;DateFilter (YYYY-DD-MM);Permissions'
                $Rows
            ) | Set-Content -Path $path -Encoding UTF8
            return $path
        }
    }

    AfterEach {
        Get-ChildItem -Path ([System.IO.Path]::GetTempPath()) -Directory -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -match '^[0-9a-fA-F-]{36}$' } |
            ForEach-Object {
                try { Remove-Item -Path $_.FullName -Recurse -Force -ErrorAction SilentlyContinue } catch {}
            }
    }

    It 'importe un mapping UTF-8 séparé par ; avec les 6 en-têtes exacts' {
        $folder = New-TestFolder
        $path = New-MappingFile -Folder $folder -Rows @(
            '\\Fileshare\RH\Paie;sharepoint;https://contoso.sharepoint.com/sites/RH/;General\\Archive\\Migration;2020-31-12;yes'
        )

        $rows = Import-FileShareMapping -Path $path

        $rows.Count | Should -Be 1
        $rows[0].TargetFolder | Should -Be 'General/Archive/Migration'
    }

    It 'accepte un mapping avec uniquement SourcePath renseigné' {
        $folder = New-TestFolder
        $path = New-MappingFile -Folder $folder -Rows @(
            '\\Fileshare\RH\Paie;;;;;'
        )

        $rows = Import-FileShareMapping -Path $path

        $rows.Count | Should -Be 1
        $rows[0].SourcePath | Should -Be '\\Fileshare\RH\Paie'
        $rows[0].TargetSPOURL | Should -Be ''
        $rows[0].TargetType | Should -Be ''
        $rows[0].Permissions | Should -Be ''
    }

    It '\\Fileshare\RH\Paie produit Fileshare_RH_Paie' {
        $folder = New-TestFolder
        $path = New-MappingFile -Folder $folder -Rows @(
            '\\Fileshare\RH\Paie;SharePoint;https://contoso.sharepoint.com/sites/RH;;2020-31-12;YES'
        )

        (Import-FileShareMapping -Path $path)[0].SourceIdentifier | Should -Be 'Fileshare_RH_Paie'
    }

    It 'génère des identifiants distincts pour RH\Paie et Compta\Paie' {
        $folder = New-TestFolder
        $path = New-MappingFile -Folder $folder -Rows @(
            '\\Fileshare\RH\Paie;SharePoint;https://contoso.sharepoint.com/sites/RH;;2020-31-12;YES',
            '\\Fileshare\Compta\Paie;SharePoint;https://contoso.sharepoint.com/sites/Finance;;2020-31-12;YES'
        )

        $rows = Import-FileShareMapping -Path $path

        $rows[0].SourceIdentifier | Should -Not -Be $rows[1].SourceIdentifier
    }

    It 'génère des identifiants distincts pour deux serveurs différents' {
        $folder = New-TestFolder
        $path = New-MappingFile -Folder $folder -Rows @(
            '\\Server1\RH\Paie;SharePoint;https://contoso.sharepoint.com/sites/RH;;2020-31-12;YES',
            '\\Server2\RH\Paie;SharePoint;https://contoso.sharepoint.com/sites/RH2;;2020-31-12;YES'
        )

        $rows = Import-FileShareMapping -Path $path

        $rows[0].SourceIdentifier | Should -Not -Be $rows[1].SourceIdentifier
    }

    It 'normalise des chemins accentués et imbriqués en identifiants valides et distincts' {
        $folder = New-TestFolder
        $path = New-MappingFile -Folder $folder -Rows @(
            '\\fs01.contoso.local\R&D\Équipe Finance\Sous-Projet A\Année-2026;SharePoint;https://contoso.sharepoint.com/sites/Finance;General;2020-31-12;YES',
            '\\fs01.contoso.local\R&D\Equipe Finance\Sous-Projet A\Annee-2026;SharePoint;https://contoso.sharepoint.com/sites/Finance-2;General;2020-31-12;YES'
        )

        $rows = Import-FileShareMapping -Path $path

        $rows[0].SourceIdentifier | Should -Match '^[A-Za-z0-9._-]+$'
        $rows[1].SourceIdentifier | Should -Match '^[A-Za-z0-9._-]+$'
        $rows[0].SourceIdentifier | Should -Not -Be $rows[1].SourceIdentifier
    }

    It 'supprime les espaces d''un segment au lieu de les remplacer par un underscore' {
        $folder = New-TestFolder
        $path = New-MappingFile -Folder $folder -Rows @(
            '\\Fileshare\RH\Paie\subfolder1\Sub Subfolder1;SharePoint;https://contoso.sharepoint.com/sites/RH;General;;YES'
        )

        $rows = Import-FileShareMapping -Path $path

        $rows[0].SourceIdentifier | Should -Be 'Fileshare_RH_Paie_subfolder1_SubSubfolder1'
    }

    It 'normalise TargetFolder avec des slashs forward' {
        $folder = New-TestFolder
        $path = New-MappingFile -Folder $folder -Rows @(
            '\\Fileshare\RH\Paie;SharePoint;https://contoso.sharepoint.com/sites/RH;General\\Archive\\Migration;2020-31-12;YES'
        )

        (Import-FileShareMapping -Path $path)[0].TargetFolder | Should -Be 'General/Archive/Migration'
    }

    It 'supprime le slash final de TargetSPOURL' {
        $folder = New-TestFolder
        $path = New-MappingFile -Folder $folder -Rows @(
            '\\Fileshare\RH\Paie;SharePoint;https://contoso.sharepoint.com/sites/RH/;;2020-31-12;YES'
        )

        (Import-FileShareMapping -Path $path)[0].TargetSPOURL | Should -Be 'https://contoso.sharepoint.com/sites/RH'
    }

    It 'accepte 2020-31-12 et rejette 2020-12-31' {
        $folder = New-TestFolder
        $ok = New-MappingFile -Folder $folder -Rows @(
            '\\Fileshare\RH\Paie;SharePoint;https://contoso.sharepoint.com/sites/RH;;2020-31-12;YES'
        )
        $ko = Join-Path $folder 'mapping-invalid.csv'
        @(
            'SourcePath;TargetType;TargetSPOURL;TargetFolder;DateFilter (YYYY-DD-MM);Permissions',
            '\\Fileshare\RH\Paie;SharePoint;https://contoso.sharepoint.com/sites/RH;;2020-12-31;YES'
        ) | Set-Content -Path $ko -Encoding UTF8

        { Import-FileShareMapping -Path $ok } | Should -Not -Throw
        (Test-FileShareMapping -Path $ko).IsValid | Should -BeFalse
    }

    It 'rejette les permissions invalides et normalise yes/oui/non/true/false' {
        $folder = New-TestFolder
        $path = New-MappingFile -Folder $folder -Rows @(
            '\\Fileshare\RH\Paie;SharePoint;https://contoso.sharepoint.com/sites/RH;;2020-31-12;yes',
            '\\Fileshare\RH\Support;SharePoint;https://contoso.sharepoint.com/sites/RH;;2020-31-12;oui',
            '\\Fileshare\RH\Archive;SharePoint;https://contoso.sharepoint.com/sites/RH;;2020-31-12;false'
        )
        $invalid = Join-Path $folder 'mapping-perm-invalid.csv'
        @(
            'SourcePath;TargetType;TargetSPOURL;TargetFolder;DateFilter (YYYY-DD-MM);Permissions',
            '\\Fileshare\RH\Paie;SharePoint;https://contoso.sharepoint.com/sites/RH;;2020-31-12;MAYBE'
        ) | Set-Content -Path $invalid -Encoding UTF8

        $rows = Import-FileShareMapping -Path $path
        $rows[0].Permissions | Should -Be 'YES'
        $rows[1].Permissions | Should -Be 'YES'
        $rows[2].Permissions | Should -Be 'NO'
        (Test-FileShareMapping -Path $invalid).IsValid | Should -BeFalse
    }

    It 'détecte un SourcePath dupliqué après normalisation' {
        $folder = New-TestFolder
        $path = New-MappingFile -Folder $folder -Rows @(
            '\\Fileshare\RH\Paie;SharePoint;https://contoso.sharepoint.com/sites/RH;;2020-31-12;YES',
            '\\Fileshare\\RH\\Paie\\;SharePoint;https://contoso.sharepoint.com/sites/RH2;;2020-31-12;YES'
        )

        $validation = Test-FileShareMapping -Path $path

        $validation.IsValid | Should -BeFalse
        ($validation.Errors.Message -join "`n") | Should -Match 'dupliqué'
    }

    It 'signale un warning en cas de périmètre imbriqué' {
        $folder = New-TestFolder
        $path = New-MappingFile -Folder $folder -Rows @(
            '\\Fileshare\RH;SharePoint;https://contoso.sharepoint.com/sites/RH;;2020-31-12;YES',
            '\\Fileshare\RH\Paie;SharePoint;https://contoso.sharepoint.com/sites/RH;Paie;2020-31-12;YES'
        )

        $validation = Test-FileShareMapping -Path $path

        $validation.IsValid | Should -BeTrue
        ($validation.Warnings.Message -join "`n") | Should -Match 'Chevauchement'
    }

    It 'signale un warning quand plusieurs sources visent la même destination' {
        $folder = New-TestFolder
        $path = New-MappingFile -Folder $folder -Rows @(
            '\\Fileshare\RH\Paie;SharePoint;https://contoso.sharepoint.com/sites/RH;Documents;2020-31-12;YES',
            '\\Fileshare\Compta\Paie;SharePoint;https://contoso.sharepoint.com/sites/RH/;Documents/;2020-31-12;YES'
        )

        $validation = Test-FileShareMapping -Path $path

        $validation.IsValid | Should -BeTrue
        ($validation.Warnings.Message -join "`n") | Should -Match 'destination'
    }

    It 'crée une arborescence de run plate avec sous-dossier par source' {
        $base = Join-Path ([System.IO.Path]::GetTempPath()) ([guid]::NewGuid().ToString())
        New-Item -ItemType Directory -Path $base -Force | Out-Null

        $run = New-AssessmentRun -Scope 'tests' -BaseOutput $base -FileSharePath 'mapping.csv'
        $sourceFolder = Get-AssessmentSourceFolder -Run $run -SourceIdentifier 'Fileshare_RH_Paie'

        Test-Path -Path $run.Path | Should -BeTrue
        (Split-Path -Leaf $sourceFolder) | Should -Be 'Fileshare_RH_Paie'
        Test-Path -Path $run.Manifest | Should -BeFalse
    }

    It 'produit des noms de fichiers contenant SourceIdentifier' {
        $base = Join-Path ([System.IO.Path]::GetTempPath()) ([guid]::NewGuid().ToString())
        New-Item -ItemType Directory -Path $base -Force | Out-Null
        $run = New-AssessmentRun -Scope 'tests' -BaseOutput $base -FileSharePath 'mapping.csv'

        $path = Get-AssessmentSourceFilePath -Run $run -SourceIdentifier 'Fileshare_RH_Paie' -Prefix 'Inventaire' -Extension 'csv'

        (Split-Path -Leaf $path) | Should -Be 'Inventaire_Fileshare_RH_Paie.csv'
    }

    It 'génère des liens relatifs cohérents dans les rapports HTML' {
        $base = Join-Path ([System.IO.Path]::GetTempPath()) ([guid]::NewGuid().ToString())
        New-Item -ItemType Directory -Path $base -Force | Out-Null
        $run = New-AssessmentRun -Scope 'tests' -BaseOutput $base -FileSharePath 'mapping.csv'
        $mappingPath = Join-Path $base 'mapping.csv'
        @(
            'SourcePath;TargetType;TargetSPOURL;TargetFolder;DateFilter (YYYY-DD-MM);Permissions',
            '\\Fileshare\RH\Paie;SharePoint;https://contoso.sharepoint.com/sites/RH;Documents;2020-31-12;YES'
        ) | Set-Content -Path $mappingPath -Encoding UTF8

        $id = 'Fileshare_RH_Paie'
        $folder = Get-AssessmentSourceFolder -Run $run -SourceIdentifier $id
        @([pscustomobject]@{ SourceIdentifier = $id; SourcePath='\\Fileshare\RH\Paie'; NombreFichiers=5; NombreDossiers=2; TailleOctets=1024; TailleLisible='1 KB'; FichiersDetail=1; DateAnalyse='2026-01-01 10:00:00' }) | Export-Csv -Path (Join-Path $folder "Synthese_${id}.csv") -Delimiter ';' -NoTypeInformation -Encoding UTF8
        @([pscustomobject]@{ NomFileShare=$id; TypeLigne='Global'; DossierNiveau1='[TOTAL]'; CheminAnalyse='\\Fileshare\RH\Paie'; NombreFichiers=5; NombreDossiers=2; TailleOctets=1024; TailleLisible='1 KB'; DateAnalyse='2026-01-01 10:00:00'; ResultatAnalyse='Inventaire global' }) | Export-Csv -Path (Join-Path $folder "Inventaire_${id}.csv") -Delimiter ';' -NoTypeInformation -Encoding UTF8
        @([pscustomobject]@{ CheminComplet='\\Fileshare\RH\Paie\a.txt'; NomFichier='a.txt'; Extension='.txt'; TailleOctets=10; TailleMB=0; Severite='WARN'; DateAnalyse='2026-01-01 10:00:00'; ResultatAnalyse='Test'; NomFileShare=$id }) | Export-Csv -Path (Join-Path $folder "Extensions_${id}.csv") -Delimiter ';' -NoTypeInformation -Encoding UTF8
        Write-EmptyCsv -Path (Join-Path $folder "Permissions_${id}.csv") -Columns @('CheminDossier','Proprietaire','Identite','TypeAcces','Droits','EstHerite','InheritanceFlags','PropagationFlags','DateAnalyse','ResultatAnalyse','NomFileShare')
        Write-EmptyCsv -Path (Join-Path $folder "CheminsLongs_${id}.csv") -Columns @('CheminComplet','Nom','TypeElement','LongueurChemin','LongueurMax','Depassement','LongueurUrlSimulee','LongueurNom','Severite','DateAnalyse','ResultatAnalyse','NomFileShare')
        Write-EmptyCsv -Path (Join-Path $folder "Doublons_${id}.csv") -Columns @('GroupeDoublonId','Hash','NombreOccurrences','TailleOctets','TailleMB','NomFichier','CheminComplet','DerniereModification','DateAnalyse','ResultatAnalyse','NomFileShare')
        Write-EmptyCsv -Path (Join-Path $folder "AccesRefuses_${id}.csv") -Columns @('NomFileShare','Chemin','TypeErreur','ExceptionType','MessageErreur','DateDetection')
        Write-EmptyCsv -Path (Join-Path $folder "InventaireDetail_${id}.csv") -Columns @('CheminComplet','CheminRelatif','DossierNiveau1','DossierParent','NomFichier','Extension','TailleOctets','TailleLisible','DateCreation','DateModification','DateDernierAcces','Attributs','EstHidden','EstSystem','EstReadOnly','EstArchive','EstReparsePoint','LongueurChemin','DateAnalyse')

        & $script:reportScriptPath -Run $run -FileShareMapping $mappingPath | Out-Null

        $report = Get-Content -Path (Join-Path $folder "Rapport_${id}.html") -Raw -Encoding UTF8
        $index = Get-Content -Path (Join-Path $run.Path 'Index_Assessment.html') -Raw -Encoding UTF8

        $report | Should -Match ('href="Inventaire_{0}\.csv"' -f $id)
        $index | Should -Match ('\.\/{0}\/Rapport_{0}\.html' -f $id)
        Test-Path -Path (Join-Path $folder "Rapport_${id}.html") | Should -BeTrue
    }

    It 'ne crée pas de sortie en WhatIf via Start-Assessment' {
        $outputPath = Join-Path -Path $script:repoRoot -ChildPath 'Scripts Assessments/Output/sample-assessment'
        if (Test-Path -Path $outputPath) {
            Remove-Item -Path $outputPath -Recurse -Force
        }

        { & $script:startAssessmentPath -WhatIf } | Should -Not -Throw
        Test-Path -Path $outputPath | Should -BeFalse
    }
}
