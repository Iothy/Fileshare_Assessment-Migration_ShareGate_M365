Describe 'FileShareAssessment packaging' {
    BeforeAll {
        $script:repoRoot = Split-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -Parent
        $script:modulePath = Join-Path -Path $script:repoRoot -ChildPath 'Scripts Assessments/FileShareAssessment/FileShareAssessment.psd1'
        $script:configPath = Join-Path -Path $script:repoRoot -ChildPath 'Scripts Assessments/Config/FileShareAssessment.json'
        $script:entryPointPath = Join-Path -Path $script:repoRoot -ChildPath 'Scripts Assessments/Start-Assessment.ps1'
        $script:outputPath = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath ("FileShareAssessment.Tests.{0}" -f [guid]::NewGuid())

        Import-Module -Name $script:modulePath -Force
    }

    AfterAll {
        if (Test-Path -LiteralPath $script:outputPath) {
            Remove-Item -LiteralPath $script:outputPath -Recurse -Force
        }
    }

    It 'importe le module sans lancer de scan' {
        Get-Module -Name FileShareAssessment | Should -Not -BeNullOrEmpty
    }

    It 'valide la configuration d''exemple' {
        $validation = Test-FileShareAssessmentConfiguration -ConfigurationPath $script:configPath

        $validation.IsValid | Should -BeTrue
        $validation.Configuration.Controls.Count | Should -Be 6
        $validation.Configuration.Assessment.MappingCsv | Should -Match 'FileShareMapping\.example\.csv$'
    }

    It 'expose le préflight sans lancer de scan' {
        Get-Command -Name Test-FileShareAssessmentPrerequisite | Should -Not -BeNullOrEmpty
    }

    It 'propose le plan sans exécuter les scripts en WhatIf' {
        $plan = Invoke-FileShareAssessment -ConfigurationPath $script:configPath -WhatIf

        $plan.Controls | Should -Contain 'inventory'
        Test-Path -Path (Join-Path -Path $script:outputPath -ChildPath 'sample-assessment') | Should -BeFalse
    }

    It 'expose un point d''entrée Start-Assessment compatible WhatIf' {
        { & $script:entryPointPath -WhatIf } | Should -Not -Throw
        Test-Path -Path (Join-Path -Path $script:outputPath -ChildPath 'sample-assessment') | Should -BeFalse
    }
}
