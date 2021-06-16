#
# Ping.ps1
#
param(
    $CrmConnectionString = "AuthType=ClientSecret;url=https://ztuat.crm11.dynamics.com/;ClientId=c2b575f9-ef65-490b-8741-800f067f2111;ClientSecret=22-OC1YovwX38F~13.K-I1gyusuh4_iUZo"
)

$ErrorActionPreference = "Stop"

Write-Verbose 'Entering Ping.ps1'

#Parameters
Write-Verbose "CrmConnectionString = $CrmConnectionString"

#Script Location
$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
Write-Verbose "Script Path: $scriptPath"

#Load XrmCIFramework
$xrmCIToolkit = $scriptPath + "\Xrm.Framework.CI.PowerShell.Cmdlets.dll"
Write-Verbose "Importing CIToolkit: $xrmCIToolkit" 
Import-Module $xrmCIToolkit
Write-Verbose "Imported CIToolkit"

#WhoAmI Check
$executingUser = Select-WhoAmI -ConnectionString $CrmConnectionString -Verbose

Write-Host "Ping Succeeded userId: " $executingUser.UserId

Write-Verbose 'Leaving Ping.ps1'