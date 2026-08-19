$script:ModuleRoot = $PSScriptRoot
$script:AssessmentRoot = Split-Path -Path $PSScriptRoot -Parent

foreach ($path in Get-ChildItem -Path (Join-Path $PSScriptRoot 'Private') -Filter '*.ps1' | Sort-Object Name) {
    . $path.FullName
}

foreach ($path in Get-ChildItem -Path (Join-Path $PSScriptRoot 'Public') -Filter '*.ps1' | Sort-Object Name) {
    . $path.FullName
}

Export-ModuleMember -Function @(
    'Import-FileShareAssessmentConfiguration',
    'Test-FileShareAssessmentConfiguration',
    'Invoke-FileShareAssessment'
)
