<#
.SYNOPSIS
Point d'entrée standardisé pour lancer un assessment complet.

.DESCRIPTION
Charge le module FileShareAssessment et exécute les scripts historiques activés
à partir d'un fichier de configuration JSON explicite.

.PARAMETER ConfigurationPath
Chemin du fichier de configuration JSON. Par défaut : .\Config\FileShareAssessment.json.

.PARAMETER Credential
Credential SMB optionnel à transmettre aux scripts historiques qui l'acceptent.

.PARAMETER Server
Nom du serveur SMB à transmettre aux scripts historiques qui l'acceptent.

.PARAMETER PassThru
Retourne le détail du run créé.

.PARAMETER PreflightOnly
Exécute uniquement les contrôles préalables de l'hôte, des sources et de la sortie.

.EXAMPLE
.\Start-Assessment.ps1 -ConfigurationPath '.\Config\FileShareAssessment.json'
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$ConfigurationPath = (Join-Path -Path $PSScriptRoot -ChildPath 'Config/FileShareAssessment.json'),

    [Parameter()]
    [System.Management.Automation.PSCredential]
    [System.Management.Automation.Credential()]
    $Credential,

    [Parameter()]
    [string]$Server,

    [Parameter()]
    [switch]$PassThru,

    [Parameter()]
    [switch]$PreflightOnly
)

$modulePath = Join-Path -Path $PSScriptRoot -ChildPath 'FileShareAssessment/FileShareAssessment.psd1'
Import-Module -Name $modulePath -Force -ErrorAction Stop

if ($PreflightOnly) {
    Test-FileShareAssessmentPrerequisite -ConfigurationPath $ConfigurationPath
    return
}

$invokeParameters = @{
    ConfigurationPath = $ConfigurationPath
    PassThru = $PassThru
}

if ($PSBoundParameters.ContainsKey('Credential')) {
    $invokeParameters.Credential = $Credential
}

if ($PSBoundParameters.ContainsKey('Server')) {
    $invokeParameters.Server = $Server
}

if ($WhatIfPreference) {
    $invokeParameters.WhatIf = $true
}

Invoke-FileShareAssessment @invokeParameters
