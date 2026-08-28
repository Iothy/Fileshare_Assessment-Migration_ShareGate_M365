Describe 'Contrôle chemins longs' {
    BeforeAll {
        $script:repoRoot = Split-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -Parent
        $script:assessmentModulePath = Join-Path -Path $script:repoRoot -ChildPath 'Scripts Assessments/Modules/PrimaGAZ.Assessment.psm1'
        $script:fileShareModulePath = Join-Path -Path $script:repoRoot -ChildPath 'Scripts Assessments/FileShareAssessment/FileShareAssessment.psd1'
        $script:pathTooLongScript = Join-Path -Path $script:repoRoot -ChildPath 'Scripts Assessments/Get-PathTooLong.ps1'

        Import-Module -Name $script:assessmentModulePath -Force
        Import-Module -Name $script:fileShareModulePath -Force

        function New-TestFolder {
            $folder = Join-Path ([System.IO.Path]::GetTempPath()) ("PathTooLong.Tests.{0}" -f [guid]::NewGuid())
            New-Item -ItemType Directory -Path $folder -Force | Out-Null
            return $folder
        }
    }

    AfterEach {
        Get-ChildItem -Path ([System.IO.Path]::GetTempPath()) -Directory -Filter 'PathTooLong.Tests.*' -ErrorAction SilentlyContinue |
            ForEach-Object {
                try { Remove-Item -LiteralPath $_.FullName -Recurse -Force -ErrorAction SilentlyContinue } catch {}
            }
    }

    It 'ignore le contrôle quand TargetSPOURL est absent sans tester l''accessibilité de la source' {
        $folder = New-TestFolder
        $mapping = Join-Path $folder 'mapping.csv'
        $output = Join-Path $folder 'out'
        @(
            'SourcePath;TargetType;TargetSPOURL;TargetFolder;DateFilter (YYYY-DD-MM);Permissions',
            '\\Fileshare\RH\Paie;;;;;'
        ) | Set-Content -Path $mapping -Encoding UTF8

        & $script:pathTooLongScript -MappingCsv $mapping -OutputPath $output | Out-Null

        $csv = Get-ChildItem -Path $output -Filter 'CheminsTropLongs_Fileshare_RH_Paie_*.csv' | Select-Object -First 1
        $row = Import-Csv -Path $csv.FullName -Delimiter ';'
        $row.StatutControle | Should -Be 'Skipped'
        $row.CompatibleSPO | Should -Be ''
        $row.CompatibleWindowsOffice | Should -Be ''
        $row.ActionRecommandee | Should -Match 'TargetSPOURL absent'
    }

    It 'exécute les contrôles SPO et Windows/Office avec cible complète' {
        $root = New-TestFolder
        $output = Join-Path $root 'out'
        $source = Join-Path $root 'source'
        New-Item -ItemType Directory -Path $source -Force | Out-Null
        $fileName = ('A' * 100) + '.txt'
        New-Item -ItemType File -Path (Join-Path $source $fileName) -Force | Out-Null
        $targetFolder = 'General/' + ('B' * 280)

        & $script:pathTooLongScript -CheminUNC $source -TargetSPOURL 'https://contoso.sharepoint.com/sites/RH' -TargetType 'Teams-Channel General' -TargetFolder $targetFolder -OutputPath $output | Out-Null

        $row = Import-Csv -Path (Get-ChildItem -Path $output -Filter 'CheminsTropLongs_*.csv' | Select-Object -First 1).FullName -Delimiter ';'
        $row.LimiteSPO | Should -Be '400'
        $row.BudgetWindowsOffice | Should -Be '160'
        $row.CompatibleSPO | Should -Be 'False'
        $row.CompatibleWindowsOffice | Should -Be 'False'
        [int]$row.LongueurCibleSPO | Should -BeGreaterThan 400
        [int]$row.LongueurRelativeWindowsOffice | Should -BeGreaterThan 160
        $row.SimulatedTargetPath | Should -Match 'https://contoso\.sharepoint\.com/sites/RH/General/'
    }

    It 'émet les fichiers et dossiers accessibles dans le pipeline' {
        Set-LogFile -Path $null
        $root = New-TestFolder
        $dir = Join-Path $root 'folder'
        $file = Join-Path $dir 'file.txt'
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
        New-Item -ItemType File -Path $file -Force | Out-Null
        $errors = New-Object 'System.Collections.Generic.List[object]'

        $items = @(Invoke-SafeRecursiveScan -RootPath $root -ErrorCollection $errors)

        @($items | Where-Object { $_ -is [System.IO.DirectoryInfo] -and $_.FullName -eq $dir }).Count | Should -Be 1
        @($items | Where-Object { $_ -is [System.IO.FileInfo] -and $_.FullName -eq $file }).Count | Should -Be 1
    }

    It 'exclut les reparse points du scan' {
        Set-LogFile -Path $null
        $root = New-TestFolder
        $target = Join-Path $root 'target'
        $link = Join-Path $root 'link'
        New-Item -ItemType Directory -Path $target -Force | Out-Null
        New-Item -ItemType SymbolicLink -Path $link -Target $target -ErrorAction Stop | Out-Null
        $errors = New-Object 'System.Collections.Generic.List[object]'

        $items = @(Invoke-SafeRecursiveScan -RootPath $root -ErrorCollection $errors)

        @($items | Where-Object { $_.FullName -eq $link }).Count | Should -Be 0
        @($errors | Where-Object { $_.TypeErreur -eq 'ReparsePointExcluded' -and $_.Chemin -eq $link }).Count | Should -Be 1
    }
}
