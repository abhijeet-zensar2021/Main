#
# ImportCMData.ps1
#

[CmdletBinding()]

param(
#[string]$crmConnectionString, #The target CRM organization connection string
[string]$CrmConnectionString = "AuthType=Office365;url=https://ztdev.crm11.dynamics.com/;UserName=Elijah@ztd365.onmicrosoft.com;Password=Zensar@1234",

[int]$crmConnectionTimeout = 10, #CRM Connection Timeout in minutes

#[string]$dataFile, #The absolute path of data zip to import
[string]$dataFile = "Data.zip",

#[string]$logsDirectory, #Optional - will place the import log in here
[string]$logsDirectory = "C:\ETA\Clients\Zensar Tech\Projects\D365 Data Migration\8.PowerShell\v1\logs",

#[string]$configurationMigrationModulePath, #The full path to the Configuration Migration PowerShell Module
[string]$configurationMigrationModulePath = "C:\ETA\Clients\Zensar Tech\Projects\D365 Data Migration\8.PowerShell\v1\cmpm",

#[string]$toolingConnectorModulePath, #The full path to the Tooling Connector PowerShell Module
[string]$toolingConnectorModulePath = "C:\ETA\Clients\Zensar Tech\Projects\D365 Data Migration\8.PowerShell\v1\tcpm",


#[int]$concurrentThreads, #Set the number of concurrent threads
[int]$concurrentThreads = 0,

#[string]$userMapFile, #User mapping file xml for on-premise deployments

#[bool]$enabledBatchMode, #Instructs CMT to import into CDS using batches
[bool]$enabledBatchMode = 1, 

[int]$batchSize = 600 #Batch size
) 

$ErrorActionPreference = "Stop"
$InformationPreference = "Continue"

Write-Verbose 'Entering ImportCMData.ps1'

#Print Parameters

Write-Verbose "crmConnectionTimeout = $crmConnectionTimeout"
Write-Verbose "dataFile = $dataFile"
Write-Verbose "logsDirectory = $logsDirectory"
Write-Verbose "concurrentThreads = $concurrentThreads"
Write-Verbose "configurationMigrationModulePath = $configurationMigrationModulePath"
Write-Verbose "toolingConnectorModulePath = $toolingConnectorModulePath"
Write-Verbose "userMapFile = $userMapFile"
Write-Verbose "enabledBatchMode = $enabledBatchMode"
Write-Verbose "batchSize = $batchSize"

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
	$connectParams.LogWriteDirectory = "$logsDirectory"
}

$CRMConn = Get-CrmConnection @connectParams -Verbose

$importParams = @{
	CrmConnection = $CRMConn
	DataFile = "$dataFile"
}
If ($logsDirectory)
{
	$importParams.LogWriteDirectory = "$logsDirectory"
}
if ($concurrentThreads -ne 0)
{
	$importParams.ConcurrentThreads = $concurrentThreads
}
if ($userMapFile)
{
	$importParams.UserMapFile = "$userMapFile"
}
if ($enabledBatchMode)
{
	$importParams.EnabledBatchMode = $enabledBatchMode
	$importParams.BatchSize = $batchSize
}

Import-CrmDataFile @importParams -EmitLogToConsole -Verbose

Write-Verbose 'Leaving ImportCMData.ps1'