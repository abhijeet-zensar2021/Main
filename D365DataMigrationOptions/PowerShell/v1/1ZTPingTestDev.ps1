#
# Ping.ps1
#
param(
    $CrmConnectionString = "AuthType=ClientSecret;url=https://ztmsdev.crm8.dynamics.com/;ClientId=50908331-b7a9-449c-b2d9-9e1fc83756a7;ClientSecret=.zO3~f1.Mwt.oN66rahdc~p6Tk5H2Avq.p"
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