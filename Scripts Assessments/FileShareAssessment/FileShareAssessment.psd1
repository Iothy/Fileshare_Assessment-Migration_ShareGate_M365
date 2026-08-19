@{
    RootModule = 'FileShareAssessment.psm1'
    ModuleVersion = '0.1.0'
    GUID = '6e6d7842-f2d1-4e36-a0a8-f18ba21cb4c6'
    Author = 'Modern Workplace Squad'
    CompanyName = 'Modern Workplace'
    Copyright = '(c) Modern Workplace Squad. All rights reserved.'
    Description = 'Socle de packaging PowerShell pour les scripts d''assessment de file shares.'
    PowerShellVersion = '5.1'
    CompatiblePSEditions = @('Desktop', 'Core')
    FunctionsToExport = @(
        'Import-FileShareAssessmentConfiguration',
        'Test-FileShareAssessmentConfiguration',
        'Invoke-FileShareAssessment'
    )
    CmdletsToExport = @()
    VariablesToExport = @()
    AliasesToExport = @()
    PrivateData = @{
        PSData = @{
            Tags = @('PowerShell', 'Assessment', 'FileShare', 'SharePoint', 'ShareGate')
            ProjectUri = 'https://github.com/Iothy/Fileshare_Assessment-Migration_ShareGate_M365'
            ReleaseNotes = 'Première phase de packaging et de standardisation des scripts historiques.'
        }
    }
}
