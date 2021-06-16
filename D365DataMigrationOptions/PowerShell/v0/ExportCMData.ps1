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
[string]$dataFile = "DevData.xml",

#[string]$schemaFile, #The absolute path to the schema file
[string]$schemaFile = "AllEntitiesSchema.xml",

#[string]$logsDirectory, #Optional - will place the import log in here
[string]$logsDirectory = "CMTExport.log",

#[string]$configurationMigrationModulePath, #The full path to the Configuration Migration PowerShell Module
[string]$configurationMigrationModulePath = "C:\ETA\Clients\Zensar Tech\Projects\D365 Data Migration\8.PowerShell\v0\CMT",

#[string]$toolingConnectorModulePath #The full path to the Tooling Connector PowerShell Module
[string]$toolingConnectorModulePath = "C:\ETA\Clients\Zensar Tech\Projects\D365 Data Migration\8.PowerShell\v0\tool" 

) 

$ErrorActionPreference = "Stop"
$InformationPreference = "Continue"

Write-Verbose 'Entering ExporCMtData.ps1'

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

Write-Verbose "Import Tooling Connector: $ToolingConnectorModulePath"
Import-module "$ToolingConnectorModulePath\Microsoft.Xrm.Tooling.CrmConnector.PowerShell.dll"

$connectParams = @{
	ConnectionString = "$CrmConnectionString"
}
if ($crmConnectionTimeout -ne 0)
{
	$connectParams.MaxCrmConnectionTimeOutMinutes = $crmConnectionTimeout
}

if ($logsDirectory)
{
	$connectParams.LogWriteDirectory = $logsDirectory
}

$CRMConn = Get-CrmConnection @connectParams -Verbose 


$exportParams = @{
	CrmConnection = "$CRMConn"
	DataFile = "$dataFile"
	SchemaFile = "$schemaFile"
}
If ($logsDirectory)
{
	$exportParams.LogWriteDirectory = $logsDirectory
}

Export-CrmDataFile @exportParams -EmitLogToConsole -Verbose

Write-Verbose 'Leaving ExporCMtData.ps1'