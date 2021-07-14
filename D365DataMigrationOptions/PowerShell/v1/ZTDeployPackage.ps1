#
# DeployPackage.ps1
#

param(
[string]$CrmConnectionString ="AuthType=ClientSecret;url=https://ztmsdev.crm8.dynamics.com/;ClientId=50908331-b7a9-449c-b2d9-9e1fc83756a7;ClientSecret=.zO3~f1.Mwt.oN66rahdc~p6Tk5H2Avq.p",
[string]$PackageName = "ZTPackageDeployment" ,
[string]$PackageDirectory = "D:\a\1\s\ZTPackageDeployment",
[string]$LogsDirectory = "D:\a\1\s\D365DataMigrationOptions\PowerShell\v1\logs",
[string]$PackageDeploymentPath ="D:\a\1\s\D365DataMigrationOptions\PowerShell\v1\tcpm2",
[string]$ToolingConnectorModulePath ="D:\a\1\s\D365DataMigrationOptions\PowerShell\v1\tcpm",
[string]$Timeout = '00:30:00', #optional timeout for Import-CrmPackage, default to 1 hour and 20 min. See https://technet.microsoft.com/en-us/library/dn756301.aspx
[int]$CrmConnectionTimeout = 2, 
[string]$RuntimePackageSettings,
[string]$UnpackFilesDirectory
)

$ErrorActionPreference = "Stop"

Write-Verbose 'Entering DeployPackage.ps1'

#Parameters
Write-Verbose "CrmConnectionString = $CrmConnectionString"
Write-Verbose "PackageName = $PackageName"
Write-Verbose "PackageDirectory = $PackageDirectory"
Write-Verbose "LogsDirectory = $LogsDirectory"
Write-Verbose "ToolingConnectorModulePath = $ToolingConnectorModulePath"
Write-Verbose "PackageDeploymentPath = $PackageDeploymentPath"
Write-Verbose "Timeout = $Timeout"
Write-Verbose "CrmConnectionTimeout = $CrmConnectionTimeout"
Write-Verbose "RuntimePackageSettings = $RuntimePackageSettings"
Write-Verbose "UnpackFilesDirectory = $UnpackFilesDirectory"

#Script Location
$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
Write-Verbose "Script Path: $scriptPath"

#Set Security Protocol
& "$scriptPath\SetTlsVersion.ps1"

#Load XRM Tooling

$crmToolingConnector = $scriptPath + "\Microsoft.Xrm.Tooling.CrmConnector.Powershell.dll"
$crmToolingDeployment = $scriptPath + "\Microsoft.Xrm.Tooling.PackageDeployment.Powershell.dll"


if($ToolingConnectorModulePath)
{
	$crmToolingConnector = $ToolingConnectorModulePath + "\Microsoft.Xrm.Tooling.CrmConnector.Powershell.dll"
}
if ($PackageDeploymentPath)
{
	$crmToolingDeployment = $PackageDeploymentPath + "\Microsoft.Xrm.Tooling.PackageDeployment.Powershell.dll"
}

Write-Verbose "Importing: $crmToolingConnector" 
Import-Module $crmToolingConnector
Write-Verbose "Imported: $crmToolingConnector"

Write-Verbose "Importing: $crmToolingDeployment" 
Import-Module $crmToolingDeployment
Write-Verbose "Imported: $crmToolingDeployment"

Write-Host "Microsoft.Xrm.Tooling.CrmConnector.Powershell.dll - Version: $([System.Diagnostics.FileVersionInfo]::GetVersionInfo($crmToolingConnector).FileVersion)"
Write-Host "Microsoft.Xrm.Tooling.PackageDeployment.Powershell.dll - Version: $([System.Diagnostics.FileVersionInfo]::GetVersionInfo($crmToolingDeployment).FileVersion)"

#Check Logs Directory
if ($LogsDirectory -eq '')
{
    $LogsDirectory = $PackageDirectory
}

#Create Connection

$CRMConn = Get-CrmConnection -ConnectionString "$CrmConnectionString" -LogWriteDirectory "$LogsDirectory" -MaxCrmConnectionTimeOutMinutes $CrmConnectionTimeout -Verbose

#Deploy Package

$PackageParams = @{
	CrmConnection = $CRMConn
	PackageDirectory = "$PackageDirectory"
	PackageName = $PackageName
	LogWriteDirectory = "$LogsDirectory"
	Timeout = $Timeout
}

if ($UnpackFilesDirectory)
{
	$PackageParams.UnpackFilesDirectory = $UnpackFilesDirectory
}

if ($RuntimePackageSettings)
{
	$PackageParams.RuntimePackageSettings = $RuntimePackageSettings
}

Import-CrmPackage @PackageParams -Verbose

Write-Verbose 'Leaving DeployPackage.ps1'
