#
# ExportCMData.ps1
#

[CmdletBinding()]

param(
#[string]$crmConnectionString, #The target CRM organization connection string
#[string]$CrmConnectionString = "AuthType=ClientSecret;url=https://ztdev.crm11.dynamics.com/;ClientId=c2b575f9-ef65-490b-8741-800f067f2111;ClientSecret=22-OC1YovwX38F~13.K-I1gyusuh4_iUZo",

 [string]$CrmConnectionString = "AuthType=Office365;url=https://ztdev.crm11.dynamics.com/;UserName=Elijah@ztd365.onmicrosoft.com;Password=Zensar@1234",

#[int]$crmConnectionTimeout, #CRM Connection Timeout in minutes
[int]$crmConnectionTimeout = 1,

#[string]$dataFile, #The absolute path of data.xml to create/update
[string]$dataFile = "DevData",

#[system.string]$dataFile = "DevData.xml",

#[string]$schemaFile, #The absolute path to the schema file
[string]$schemaFile = "AllEntitiesSchema.xml",

#[string]$logsDirectory, #Optional - will place the import log in here
[string]$logsDirectory = "CMTExport.log",

#$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition,
#Write-Host "Script Path1: $scriptPath",
#Write-Host "Script Path2: $scriptPath",
#Write-Verbose "Script Path3: $scriptPath",

#[string]$configurationMigrationModulePath, #The full path to the Configuration Migration PowerShell Module
#[string]$configurationMigrationModulePath = $scriptPath + "\CMT",
[string]$configurationMigrationModulePath = "D:\a\1\s\D365DataMigrationOptions\PowerShell\v0\cmt",



#[string]$toolingConnectorModulePath #The full path to the Tooling Connector PowerShell Module
#[string]$toolingConnectorModulePath = $scriptPath + "\tool" 
[string]$toolingConnectorModulePath = "D:\a\1\s\D365DataMigrationOptions\PowerShell\v0\tool" 

) 

$ErrorActionPreference = "Stop"
$InformationPreference = "Continue"

Write-Verbose 'Entering ExporCMtData.ps1'
Write-Host 'Entering ExporCMtData.ps1'

#Print Parameters

Write-Verbose "CrmConnectionString = $CrmConnectionString"
Write-Verbose "dataFile = $dataFile"
Write-Verbose "schemaFile = $schemaFile"
Write-Verbose "logsDirectory = $logsDirectory"
Write-Verbose "crmConnectionTimeout = $crmConnectionTimeout"
Write-Verbose "configurationMigrationModulePath = $configurationMigrationModulePath"
Write-Verbose "toolingConnectorModulePath = $toolingConnectorModulePath"

#Script Location
$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
Write-Verbose "Script Path: $scriptPath"

#Load PS Module
Write-Verbose "Importing Configuration Migration: $PowerAppsCheckerPath"
Import-module "$ConfigurationMigrationModulePath\Microsoft.Xrm.Tooling.ConfigurationMigration.psd1"
#cls
Write-Verbose "Import Tooling Connector: $ToolingConnectorModulePath"
Import-module "$ToolingConnectorModulePath\Microsoft.Xrm.Tooling.CrmConnector.PowerShell.dll"
Write-Host "Import Tooling Connector: $ToolingConnectorModulePath"

$connectParams = $CrmConnectionString
Write-Host "Connection String: $CrmConnectionString"


$connectParams = @{
	ConnectionString = "$CrmConnectionString"
}
Write-Host "Connection Params: $connectParams"

if ($crmConnectionTimeout -ne 0)
{
	$connectParams.MaxCrmConnectionTimeOutMinutes = $crmConnectionTimeout
}

Write-Host "Connection Params2: $connectParams"

if ($logsDirectory)
{
	$connectParams.LogWriteDirectory = $logsDirectory
}
Write-Host "Connection Params3: $connectParams"

$CRMConn = Get-CrmConnection @connectParams -Verbose
Write-Host "CRMConnection: $CRMConn"


$exportParams = @{
	CrmConnection = $CRMConn
	DataFile = "$dataFile"
	SchemaFile = "$schemaFile"
}
Write-Host "Export Params: $exportParams"

If ($logsDirectory)
{
	$exportParams.LogWriteDirectory = $logsDirectory
}
Write-Host "Export Params Log: $exportParams.LogWriteDirectory"

Export-CrmDataFile @exportParams -EmitLogToConsole -Verbose
#Export-CrmDataFile @exportParams  -Verbose
#Export-CrmDataFile -CrmConnection $CRMConn -SchemaFile $schemaFile -DataFile $dataFile

Write-Verbose 'Leaving ExporCMtData.ps1'
Write-Host 'Leaving ExporCMtData.ps1'