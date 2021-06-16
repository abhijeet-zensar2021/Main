#
# Module manifest for module 'Microsoft.Xrm.Tooling.ConfigurationMigration'
#

@{

# Script module or binary module file associated with this manifest.
RootModule = 'Microsoft.Xrm.Tooling.ConfigurationMigration.psm1'

# Version number of this module.
ModuleVersion = '1.0.0.60'

# ID used to uniquely identify this module
GUID = '162392C2-EDE9-4754-B5A3-7DE42E5C768C'

# Author of this module
Author = 'Microsoft Common Data Service Team'

# Company or vendor of this module
CompanyName = 'Microsoft'

# Copyright statement for this module
Copyright = '© 2020 Microsoft Corporation. All rights reserved'

# Description of the functionality provided by this module
Description = 'PowerShell Module for Microsoft Dynamics CRM Configuration Migration Utility'

# Minimum version of the Windows PowerShell engine required by this module
#PowerShellVersion = '3.0'

# Name of the Windows PowerShell host required by this module
#PowerShellHostName = ''

# Minimum version of the Windows PowerShell host required by this module
#PowerShellHostVersion = ''

# Minimum version of the .NET Framework required by this module
#DotNetFrameworkVersion = '4.0'

# Minimum version of the common language runtime (CLR) required by this module
#CLRVersion = '4.0'

# Processor architecture (None, X86, Amd64) required by this module
#ProcessorArchitecture = ''

# Modules that must be imported into the global environment prior to importing this module
#RequiredModules = 

# Assemblies that must be loaded prior to importing this module
#RequiredAssemblies = @()

# Script files (.ps1) that are run in the caller's environment prior to importing this module.
#ScriptsToProcess = @()

# Type files (.ps1xml) to be loaded when importing this module
#TypesToProcess = @()

# Format files (.ps1xml) to be loaded when importing this module
#FormatsToProcess = @()

# Modules to import as nested modules of the module specified in RootModule/ModuleToProcess
NestedModules = @(
	'Microsoft.Xrm.Tooling.ConfigurationMigration.dll'
	)

# Functions to export from this module
FunctionsToExport = @()

# Cmdlets to export from this module
CmdletsToExport = @(
	'Import-CrmDataFile'
	'Export-CrmDataFile'
	)

# Variables to export from this module
VariablesToExport = @(
	)

# Aliases to export from this module
AliasesToExport = @(
	)

# List of all modules packaged with this module.
ModuleList = @(
	'Microsoft.Xrm.Tooling.ConfigurationMigration'
	)

# List of all files packaged with this module
#FileList = @()

# Private data to pass to the module specified in RootModule/ModuleToProcess. This may also contain a PSData hashtable with additional module metadata used by PowerShell.
PrivateData = @{

    PSData = @{

        # Tags applied to this module. These help with module discovery in online galleries.
        # Tags = @()

        # A URL to the license for this module.
        LicenseUri = 'https://go.microsoft.com/fwlink/?linkid=2108407'

        # A URL to the main website for this project.
        #ProjectUri = 'https://docs.microsoft.com/powershell/module/microsoft.xrm.tooling.crmconnector.powershell/?view=dynamics365ce-ps'

        # A URL to an icon representing this module.
        IconUri = 'https://connectoricons-prod.azureedge.net/powerappsforappmakers/icon_1.0.1056.1255.png'


        # ReleaseNotes of this module
        ReleaseNotes = '

1.0.0.60
Added support for Updating Dates on insert to shift dates to Day of insert.
Added option to disable telemetry
Added support for inmemory log capture for package deployer and CMT
Corrected an issue that would cause CMT to hang when executing prefetch. this situation was caused by very large record counts in entities that were being prefetched.
Corrected an issue with CMT lookup resolution that would prefer the display name over the ID of the target of the lookup.

1.0.0.53:
Corrected an issue with CMT where a failure to import would occur if, during the decision process for update vs insert, an existing record was not found by ID and the Name field of the record was null
 This would only impact default search patterns.  Entities that defined custom searchs would not be impacted.

1.0.0.52:
Fixed an issue with Batch mode import that is caused by invalid guids being hand edited into datafiles.
Fixed an issue where CMT would prefer ID/Name over comparison keys when precaching data for comparison.

1.0.0.51:
Fix in Data import to compensate for Import files created during a window of time where new virtual fields generated by CDS were incorrectly identified as valid fields.
This issue would cause schema validation to fail during import of complex data for some types of attributeOf type fields. 
Up-taking fixes for Sovereign cloud discovery connectivity for login controls

1.0.0.49:
Fixed an issue where data import could hang without error
Added new optional parameter called PrefetchRecordLimitCount ( defaults to 4000 ) this controls the number of records that are prefetched when doing pre-validation during import.
Fix for an issue where importing business units that included a lookup for their own team (looped back to the entity) would not be populated on update.
Fixed a general issue were related LookUp IDs that were discovered in the target system were not properly linked.
Behavior Change: when using BatchMode, Create will always include the ID of the record from the source system, even if not explicitly included as a field in the exported system.

3.3.0.891:
Fix for an issue reported where related objects were not being correctly identified by name and thus not correctly linked to importing records
This only impacted imports where the related entity did not have comparison keys

1.0.0.35:
Fix for Entity model changes in CDS to support filtering out non importable fields introduced by new entity types.

1.0.0.34:
Fix for a fail to import bug that can occur if an incoming data file has either a state or status code but not both
Fix for a non-failing data error that can occur when creating pre-parse data for import. 
Fix for activating Duplicate Detection rules properly post when present.

1.0.0.32:
Speculative fix for Concurrency issue with Batch and Threaded Mode import. 
Modified threaded import to improve performance and throughput. 

1.0.0.29:
Fixed a bug where status of a record would not be updated for records in an active state. 
Fixed a bug in batch mode when the client is provided only 1 connection to work with.
Optimized updates for state/status events to prevent unnecessary updates to Cds. (only affects version 9.0 +)
Reduced chattiness of logs, logs will now only emit once for the same type per thread. 

1.0.0.26:
Fixed an issue preventing logs from being generated correctly,  Logs should once again write to the log directory

1.0.0.21:
Added new cmdlet switch called EmitLogToConsole.  When set, the log will be written to the console of the powershell session in addition to logs.  Intended to used in conjunction with the -verbose parameter 
Increased default timeout for Export-CrmData cmdlet. 
Added 2 new parameters to Import-CrmData
	EnabledBatchMode – instructs CMT to import into CDS using batches ( default size of 600 records ).  Note this does work in conjunction with multiple connections,  thus 2 connections creates 2 batches of 600 records. 
	BatchSize -  Allows you to override the default batch size ( 600 records ).  
Fixed a bug with pre-processing in CMT that would cause the system to runout of memory when importing to entities that had large numbers of records. 


1.0.0.16:
    Inital release for Configuration Migration Tool Powershell. 
        '

    } # End of PSData hashtable

} # End of PrivateData hashtable

# HelpInfo URI of this module
#HelpInfoURI = ''

# Default prefix for commands exported from this module. Override the default prefix using Import-Module -Prefix.
#DefaultCommandPrefix = ''

}

# SIG # Begin signature block
# MIIjkQYJKoZIhvcNAQcCoIIjgjCCI34CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCpmpsAxqp+T/zM
# ASWOg0G78X3TFB9eSNk2oxfXMCxBoaCCDYEwggX/MIID56ADAgECAhMzAAAB32vw
# LpKnSrTQAAAAAAHfMA0GCSqGSIb3DQEBCwUAMH4xCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25p
# bmcgUENBIDIwMTEwHhcNMjAxMjE1MjEzMTQ1WhcNMjExMjAyMjEzMTQ1WjB0MQsw
# CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9u
# ZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMR4wHAYDVQQDExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24wggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIB
# AQC2uxlZEACjqfHkuFyoCwfL25ofI9DZWKt4wEj3JBQ48GPt1UsDv834CcoUUPMn
# s/6CtPoaQ4Thy/kbOOg/zJAnrJeiMQqRe2Lsdb/NSI2gXXX9lad1/yPUDOXo4GNw
# PjXq1JZi+HZV91bUr6ZjzePj1g+bepsqd/HC1XScj0fT3aAxLRykJSzExEBmU9eS
# yuOwUuq+CriudQtWGMdJU650v/KmzfM46Y6lo/MCnnpvz3zEL7PMdUdwqj/nYhGG
# 3UVILxX7tAdMbz7LN+6WOIpT1A41rwaoOVnv+8Ua94HwhjZmu1S73yeV7RZZNxoh
# EegJi9YYssXa7UZUUkCCA+KnAgMBAAGjggF+MIIBejAfBgNVHSUEGDAWBgorBgEE
# AYI3TAgBBggrBgEFBQcDAzAdBgNVHQ4EFgQUOPbML8IdkNGtCfMmVPtvI6VZ8+Mw
# UAYDVR0RBEkwR6RFMEMxKTAnBgNVBAsTIE1pY3Jvc29mdCBPcGVyYXRpb25zIFB1
# ZXJ0byBSaWNvMRYwFAYDVQQFEw0yMzAwMTIrNDYzMDA5MB8GA1UdIwQYMBaAFEhu
# ZOVQBdOCqhc3NyK1bajKdQKVMFQGA1UdHwRNMEswSaBHoEWGQ2h0dHA6Ly93d3cu
# bWljcm9zb2Z0LmNvbS9wa2lvcHMvY3JsL01pY0NvZFNpZ1BDQTIwMTFfMjAxMS0w
# Ny0wOC5jcmwwYQYIKwYBBQUHAQEEVTBTMFEGCCsGAQUFBzAChkVodHRwOi8vd3d3
# Lm1pY3Jvc29mdC5jb20vcGtpb3BzL2NlcnRzL01pY0NvZFNpZ1BDQTIwMTFfMjAx
# MS0wNy0wOC5jcnQwDAYDVR0TAQH/BAIwADANBgkqhkiG9w0BAQsFAAOCAgEAnnqH
# tDyYUFaVAkvAK0eqq6nhoL95SZQu3RnpZ7tdQ89QR3++7A+4hrr7V4xxmkB5BObS
# 0YK+MALE02atjwWgPdpYQ68WdLGroJZHkbZdgERG+7tETFl3aKF4KpoSaGOskZXp
# TPnCaMo2PXoAMVMGpsQEQswimZq3IQ3nRQfBlJ0PoMMcN/+Pks8ZTL1BoPYsJpok
# t6cql59q6CypZYIwgyJ892HpttybHKg1ZtQLUlSXccRMlugPgEcNZJagPEgPYni4
# b11snjRAgf0dyQ0zI9aLXqTxWUU5pCIFiPT0b2wsxzRqCtyGqpkGM8P9GazO8eao
# mVItCYBcJSByBx/pS0cSYwBBHAZxJODUqxSXoSGDvmTfqUJXntnWkL4okok1FiCD
# Z4jpyXOQunb6egIXvkgQ7jb2uO26Ow0m8RwleDvhOMrnHsupiOPbozKroSa6paFt
# VSh89abUSooR8QdZciemmoFhcWkEwFg4spzvYNP4nIs193261WyTaRMZoceGun7G
# CT2Rl653uUj+F+g94c63AhzSq4khdL4HlFIP2ePv29smfUnHtGq6yYFDLnT0q/Y+
# Di3jwloF8EWkkHRtSuXlFUbTmwr/lDDgbpZiKhLS7CBTDj32I0L5i532+uHczw82
# oZDmYmYmIUSMbZOgS65h797rj5JJ6OkeEUJoAVwwggd6MIIFYqADAgECAgphDpDS
# AAAAAAADMA0GCSqGSIb3DQEBCwUAMIGIMQswCQYDVQQGEwJVUzETMBEGA1UECBMK
# V2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0
# IENvcnBvcmF0aW9uMTIwMAYDVQQDEylNaWNyb3NvZnQgUm9vdCBDZXJ0aWZpY2F0
# ZSBBdXRob3JpdHkgMjAxMTAeFw0xMTA3MDgyMDU5MDlaFw0yNjA3MDgyMTA5MDla
# MH4xCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdS
# ZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMT
# H01pY3Jvc29mdCBDb2RlIFNpZ25pbmcgUENBIDIwMTEwggIiMA0GCSqGSIb3DQEB
# AQUAA4ICDwAwggIKAoICAQCr8PpyEBwurdhuqoIQTTS68rZYIZ9CGypr6VpQqrgG
# OBoESbp/wwwe3TdrxhLYC/A4wpkGsMg51QEUMULTiQ15ZId+lGAkbK+eSZzpaF7S
# 35tTsgosw6/ZqSuuegmv15ZZymAaBelmdugyUiYSL+erCFDPs0S3XdjELgN1q2jz
# y23zOlyhFvRGuuA4ZKxuZDV4pqBjDy3TQJP4494HDdVceaVJKecNvqATd76UPe/7
# 4ytaEB9NViiienLgEjq3SV7Y7e1DkYPZe7J7hhvZPrGMXeiJT4Qa8qEvWeSQOy2u
# M1jFtz7+MtOzAz2xsq+SOH7SnYAs9U5WkSE1JcM5bmR/U7qcD60ZI4TL9LoDho33
# X/DQUr+MlIe8wCF0JV8YKLbMJyg4JZg5SjbPfLGSrhwjp6lm7GEfauEoSZ1fiOIl
# XdMhSz5SxLVXPyQD8NF6Wy/VI+NwXQ9RRnez+ADhvKwCgl/bwBWzvRvUVUvnOaEP
# 6SNJvBi4RHxF5MHDcnrgcuck379GmcXvwhxX24ON7E1JMKerjt/sW5+v/N2wZuLB
# l4F77dbtS+dJKacTKKanfWeA5opieF+yL4TXV5xcv3coKPHtbcMojyyPQDdPweGF
# RInECUzF1KVDL3SV9274eCBYLBNdYJWaPk8zhNqwiBfenk70lrC8RqBsmNLg1oiM
# CwIDAQABo4IB7TCCAekwEAYJKwYBBAGCNxUBBAMCAQAwHQYDVR0OBBYEFEhuZOVQ
# BdOCqhc3NyK1bajKdQKVMBkGCSsGAQQBgjcUAgQMHgoAUwB1AGIAQwBBMAsGA1Ud
# DwQEAwIBhjAPBgNVHRMBAf8EBTADAQH/MB8GA1UdIwQYMBaAFHItOgIxkEO5FAVO
# 4eqnxzHRI4k0MFoGA1UdHwRTMFEwT6BNoEuGSWh0dHA6Ly9jcmwubWljcm9zb2Z0
# LmNvbS9wa2kvY3JsL3Byb2R1Y3RzL01pY1Jvb0NlckF1dDIwMTFfMjAxMV8wM18y
# Mi5jcmwwXgYIKwYBBQUHAQEEUjBQME4GCCsGAQUFBzAChkJodHRwOi8vd3d3Lm1p
# Y3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY1Jvb0NlckF1dDIwMTFfMjAxMV8wM18y
# Mi5jcnQwgZ8GA1UdIASBlzCBlDCBkQYJKwYBBAGCNy4DMIGDMD8GCCsGAQUFBwIB
# FjNodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpb3BzL2RvY3MvcHJpbWFyeWNw
# cy5odG0wQAYIKwYBBQUHAgIwNB4yIB0ATABlAGcAYQBsAF8AcABvAGwAaQBjAHkA
# XwBzAHQAYQB0AGUAbQBlAG4AdAAuIB0wDQYJKoZIhvcNAQELBQADggIBAGfyhqWY
# 4FR5Gi7T2HRnIpsLlhHhY5KZQpZ90nkMkMFlXy4sPvjDctFtg/6+P+gKyju/R6mj
# 82nbY78iNaWXXWWEkH2LRlBV2AySfNIaSxzzPEKLUtCw/WvjPgcuKZvmPRul1LUd
# d5Q54ulkyUQ9eHoj8xN9ppB0g430yyYCRirCihC7pKkFDJvtaPpoLpWgKj8qa1hJ
# Yx8JaW5amJbkg/TAj/NGK978O9C9Ne9uJa7lryft0N3zDq+ZKJeYTQ49C/IIidYf
# wzIY4vDFLc5bnrRJOQrGCsLGra7lstnbFYhRRVg4MnEnGn+x9Cf43iw6IGmYslmJ
# aG5vp7d0w0AFBqYBKig+gj8TTWYLwLNN9eGPfxxvFX1Fp3blQCplo8NdUmKGwx1j
# NpeG39rz+PIWoZon4c2ll9DuXWNB41sHnIc+BncG0QaxdR8UvmFhtfDcxhsEvt9B
# xw4o7t5lL+yX9qFcltgA1qFGvVnzl6UJS0gQmYAf0AApxbGbpT9Fdx41xtKiop96
# eiL6SJUfq/tHI4D1nvi/a7dLl+LrdXga7Oo3mXkYS//WsyNodeav+vyL6wuA6mk7
# r/ww7QRMjt/fdW1jkT3RnVZOT7+AVyKheBEyIXrvQQqxP/uozKRdwaGIm1dxVk5I
# RcBCyZt2WwqASGv9eZ/BvW1taslScxMNelDNMYIVZjCCFWICAQEwgZUwfjELMAkG
# A1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQx
# HjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEoMCYGA1UEAxMfTWljcm9z
# b2Z0IENvZGUgU2lnbmluZyBQQ0EgMjAxMQITMwAAAd9r8C6Sp0q00AAAAAAB3zAN
# BglghkgBZQMEAgEFAKCBoDAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgor
# BgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAvBgkqhkiG9w0BCQQxIgQgaDG/VkAm
# TTjWisakLhlF75RAydZX+oewajY9qhVeVPQwNAYKKwYBBAGCNwIBDDEmMCSgEoAQ
# AFQAZQBzAHQAUwBpAGcAbqEOgAxodHRwOi8vdGVzdCAwDQYJKoZIhvcNAQEBBQAE
# ggEAcHh8RCkSMfSfJsXF9DeA44lVEY7cnX8I0yDHt0HrHEW+adRR5xpUDBpvtwck
# 3GFIXzfKwYBcZ4skLwfOQmryp/cLj+g1tSHYHTz7ykDzhnjCn1p+NKPBRRDDE7Ky
# frynvnrRGurCtMIIMzAoEogkTnvGCUqrfeecZjjQEnIGVk5h15kxRsNnRcLKWEvO
# +Jxc4I2ma2/2uoFV1CuMvc4UoKAB0Yw/3AyamQp+B5utjCdCGZEsZhY7q5/iGsmJ
# O3jAczzdzTGZDC9DZBd1WYDyWe8HurOuU+5ckgukiv0y/gPVowNN6DAyHSDiXyQD
# S/y8/3TKDK0NirP/dkDgGijqYaGCEv4wghL6BgorBgEEAYI3AwMBMYIS6jCCEuYG
# CSqGSIb3DQEHAqCCEtcwghLTAgEDMQ8wDQYJYIZIAWUDBAIBBQAwggFZBgsqhkiG
# 9w0BCRABBKCCAUgEggFEMIIBQAIBAQYKKwYBBAGEWQoDATAxMA0GCWCGSAFlAwQC
# AQUABCAgt+MG1RPL+gYL4JzRCeXte1NHr7G8TDAvS55amCyVzAIGYIrORrM0GBMy
# MDIxMDUyMTAxNDI0MS4zNTFaMASAAgH0oIHYpIHVMIHSMQswCQYDVQQGEwJVUzET
# MBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMV
# TWljcm9zb2Z0IENvcnBvcmF0aW9uMS0wKwYDVQQLEyRNaWNyb3NvZnQgSXJlbGFu
# ZCBPcGVyYXRpb25zIExpbWl0ZWQxJjAkBgNVBAsTHVRoYWxlcyBUU1MgRVNOOjhE
# NDEtNEJGNy1CM0I3MSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBTZXJ2
# aWNloIIOTTCCBPkwggPhoAMCAQICEzMAAAE6jY0x93dJScIAAAAAATowDQYJKoZI
# hvcNAQELBQAwfDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAO
# BgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEm
# MCQGA1UEAxMdTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBIDIwMTAwHhcNMjAxMDE1
# MTcyODIyWhcNMjIwMTEyMTcyODIyWjCB0jELMAkGA1UEBhMCVVMxEzARBgNVBAgT
# Cldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29m
# dCBDb3Jwb3JhdGlvbjEtMCsGA1UECxMkTWljcm9zb2Z0IElyZWxhbmQgT3BlcmF0
# aW9ucyBMaW1pdGVkMSYwJAYDVQQLEx1UaGFsZXMgVFNTIEVTTjo4RDQxLTRCRjct
# QjNCNzElMCMGA1UEAxMcTWljcm9zb2Z0IFRpbWUtU3RhbXAgU2VydmljZTCCASIw
# DQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBAM5fJOdfD5c/CUyF2J/zvTmpKnFq
# nSfGVyYQJLPYciwfPgDVu2Z4G9be4sO05oHqKDsDJ24QkEB9eFMFnfIUBVnsSqla
# XCBZ2N29efj2JoFXkOFyn1idzzzI04l7u93qVx4/V1+5oUtiFasBLwzrnTiN8kC/
# A/DjhtuG/tdMwwxMjecL2eQNVnL5dnkIQhutRnhzWl71zpaE0G5VRWzFjphr4h74
# EZUYUY4DZrr9Sdsun0BbYhMNKXVygkBuvMUJQYXLuC5/m2C+B6Hk9pq7ZKRxMg2f
# n66/imv81Az9wbHB5CnZDRjjg37yXVh2ldFs69cJz7lILTm35wTtKEoyKQUCAwEA
# AaOCARswggEXMB0GA1UdDgQWBBQUVqCfFl8Sa6bJbxx1rK0JhUXzyzAfBgNVHSME
# GDAWgBTVYzpcijGQ80N7fEYbxTNoWoVtVTBWBgNVHR8ETzBNMEugSaBHhkVodHRw
# Oi8vY3JsLm1pY3Jvc29mdC5jb20vcGtpL2NybC9wcm9kdWN0cy9NaWNUaW1TdGFQ
# Q0FfMjAxMC0wNy0wMS5jcmwwWgYIKwYBBQUHAQEETjBMMEoGCCsGAQUFBzAChj5o
# dHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY1RpbVN0YVBDQV8y
# MDEwLTA3LTAxLmNydDAMBgNVHRMBAf8EAjAAMBMGA1UdJQQMMAoGCCsGAQUFBwMI
# MA0GCSqGSIb3DQEBCwUAA4IBAQBeN+Q+pAEto3gCfARtSk5uQM9I6W3Y6akrzC7e
# 3zla2gBB5XYJNOASnE5oc42017I8IAjDAkvo1+E6KotHJ83EqA9i6YWQPfA13h+p
# UR+6kHV2x4MdP726tGMlZJbUkzL9Y88yO+WeXYab3OFMJ+BW+ggmhdv/q+0a9I/t
# gxgsJfGlqr0Ks4Qif6O3DaGFHyeFkgh7NTrI4DJlk5oJbjyMbd5FfIi2ZdiGYMq+
# XDTojYk62hQ6u4cSNeqA7KFtBECglfOJ1dcFklntUogCiMK+QbamMCfoZaHD8Oan
# oVmwYky57XC68CnQG8LYRa9Qx390WMBaA88jbSkgf0ybAdmzMIIGcTCCBFmgAwIB
# AgIKYQmBKgAAAAAAAjANBgkqhkiG9w0BAQsFADCBiDELMAkGA1UEBhMCVVMxEzAR
# BgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1p
# Y3Jvc29mdCBDb3Jwb3JhdGlvbjEyMDAGA1UEAxMpTWljcm9zb2Z0IFJvb3QgQ2Vy
# dGlmaWNhdGUgQXV0aG9yaXR5IDIwMTAwHhcNMTAwNzAxMjEzNjU1WhcNMjUwNzAx
# MjE0NjU1WjB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSYw
# JAYDVQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMDCCASIwDQYJKoZI
# hvcNAQEBBQADggEPADCCAQoCggEBAKkdDbx3EYo6IOz8E5f1+n9plGt0VBDVpQoA
# goX77XxoSyxfxcPlYcJ2tz5mK1vwFVMnBDEfQRsalR3OCROOfGEwWbEwRA/xYIiE
# VEMM1024OAizQt2TrNZzMFcmgqNFDdDq9UeBzb8kYDJYYEbyWEeGMoQedGFnkV+B
# VLHPk0ySwcSmXdFhE24oxhr5hoC732H8RsEnHSRnEnIaIYqvS2SJUGKxXf13Hz3w
# V3WsvYpCTUBR0Q+cBj5nf/VmwAOWRH7v0Ev9buWayrGo8noqCjHw2k4GkbaICDXo
# eByw6ZnNPOcvRLqn9NxkvaQBwSAJk3jN/LzAyURdXhacAQVPIk0CAwEAAaOCAeYw
# ggHiMBAGCSsGAQQBgjcVAQQDAgEAMB0GA1UdDgQWBBTVYzpcijGQ80N7fEYbxTNo
# WoVtVTAZBgkrBgEEAYI3FAIEDB4KAFMAdQBiAEMAQTALBgNVHQ8EBAMCAYYwDwYD
# VR0TAQH/BAUwAwEB/zAfBgNVHSMEGDAWgBTV9lbLj+iiXGJo0T2UkFvXzpoYxDBW
# BgNVHR8ETzBNMEugSaBHhkVodHRwOi8vY3JsLm1pY3Jvc29mdC5jb20vcGtpL2Ny
# bC9wcm9kdWN0cy9NaWNSb29DZXJBdXRfMjAxMC0wNi0yMy5jcmwwWgYIKwYBBQUH
# AQEETjBMMEoGCCsGAQUFBzAChj5odHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtp
# L2NlcnRzL01pY1Jvb0NlckF1dF8yMDEwLTA2LTIzLmNydDCBoAYDVR0gAQH/BIGV
# MIGSMIGPBgkrBgEEAYI3LgMwgYEwPQYIKwYBBQUHAgEWMWh0dHA6Ly93d3cubWlj
# cm9zb2Z0LmNvbS9QS0kvZG9jcy9DUFMvZGVmYXVsdC5odG0wQAYIKwYBBQUHAgIw
# NB4yIB0ATABlAGcAYQBsAF8AUABvAGwAaQBjAHkAXwBTAHQAYQB0AGUAbQBlAG4A
# dAAuIB0wDQYJKoZIhvcNAQELBQADggIBAAfmiFEN4sbgmD+BcQM9naOhIW+z66bM
# 9TG+zwXiqf76V20ZMLPCxWbJat/15/B4vceoniXj+bzta1RXCCtRgkQS+7lTjMz0
# YBKKdsxAQEGb3FwX/1z5Xhc1mCRWS3TvQhDIr79/xn/yN31aPxzymXlKkVIArzgP
# F/UveYFl2am1a+THzvbKegBvSzBEJCI8z+0DpZaPWSm8tv0E4XCfMkon/VWvL/62
# 5Y4zu2JfmttXQOnxzplmkIz/amJ/3cVKC5Em4jnsGUpxY517IW3DnKOiPPp/fZZq
# kHimbdLhnPkd/DjYlPTGpQqWhqS9nhquBEKDuLWAmyI4ILUl5WTs9/S/fmNZJQ96
# LjlXdqJxqgaKD4kWumGnEcua2A5HmoDF0M2n0O99g/DhO3EJ3110mCIIYdqwUB5v
# vfHhAN/nMQekkzr3ZUd46PioSKv33nJ+YWtvd6mBy6cJrDm77MbL2IK0cs0d9LiF
# AR6A+xuJKlQ5slvayA1VmXqHczsI5pgt6o3gMy4SKfXAL1QnIffIrE7aKLixqduW
# sqdCosnPGUFN4Ib5KpqjEWYw07t0MkvfY3v1mYovG8chr1m1rtxEPJdQcdeh0sVV
# 42neV8HR3jDA/czmTfsNv11P6Z0eGTgvvM9YBS7vDaBQNdrvCScc1bN+NR4Iuto2
# 29Nfj950iEkSoYIC1zCCAkACAQEwggEAoYHYpIHVMIHSMQswCQYDVQQGEwJVUzET
# MBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMV
# TWljcm9zb2Z0IENvcnBvcmF0aW9uMS0wKwYDVQQLEyRNaWNyb3NvZnQgSXJlbGFu
# ZCBPcGVyYXRpb25zIExpbWl0ZWQxJjAkBgNVBAsTHVRoYWxlcyBUU1MgRVNOOjhE
# NDEtNEJGNy1CM0I3MSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBTZXJ2
# aWNloiMKAQEwBwYFKw4DAhoDFQAHJZHZ9Y9YF4Hcr2I+NQfK5DWeCKCBgzCBgKR+
# MHwxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdS
# ZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJjAkBgNVBAMT
# HU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwMA0GCSqGSIb3DQEBBQUAAgUA
# 5FD7bzAiGA8yMDIxMDUyMDIzMTQ1NVoYDzIwMjEwNTIxMjMxNDU1WjB3MD0GCisG
# AQQBhFkKBAExLzAtMAoCBQDkUPtvAgEAMAoCAQACAh8aAgH/MAcCAQACAhEdMAoC
# BQDkUkzvAgEAMDYGCisGAQQBhFkKBAIxKDAmMAwGCisGAQQBhFkKAwKgCjAIAgEA
# AgMHoSChCjAIAgEAAgMBhqAwDQYJKoZIhvcNAQEFBQADgYEAZ6ShDdv/RYmv1+e4
# bqR1YxT+jgE4jMu8Naf6l3jivVpc/hcram2SjvSe+EEdhAFt07FOfRv4H/zgRmhG
# pVOjoeTGJKohPXbcAEKQop+8SF1gMzzlsewBkWSRZMNrn56yclVj8Uj05THc9hX/
# Mvtv90WJOVk9Sy7pfkzNt9lsHskxggMNMIIDCQIBATCBkzB8MQswCQYDVQQGEwJV
# UzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UE
# ChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQgVGlt
# ZS1TdGFtcCBQQ0EgMjAxMAITMwAAATqNjTH3d0lJwgAAAAABOjANBglghkgBZQME
# AgEFAKCCAUowGgYJKoZIhvcNAQkDMQ0GCyqGSIb3DQEJEAEEMC8GCSqGSIb3DQEJ
# BDEiBCAWEaVMoI4eDcehT3pRynv/FJffp66M5Yubr7Foqb0IBTCB+gYLKoZIhvcN
# AQkQAi8xgeowgecwgeQwgb0EIJ+v0IQHqSxf+wXbL37vBjk/ooS/XOHIKOTX9WlD
# fLRtMIGYMIGApH4wfDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24x
# EDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlv
# bjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBIDIwMTACEzMAAAE6
# jY0x93dJScIAAAAAATowIgQga57yHOVehFhXn9qzgoxXEDJfli2D7wFmdHTND3QG
# QUcwDQYJKoZIhvcNAQELBQAEggEAizo1dtl4/GSkxBXk+JSB46Lkf4tGgg0ToILI
# nozcCH5YTpPvFKZEZZaAZNt0+yxKliCTb9FZnnVY8temIowXvhdtURz/a628oDSh
# rx/KcbNWCALUWbmd5XSeQS+c35OzKbq7pTqehSJSH0KD/HTgZhh6BjvbkxTGNQhl
# XzJtZGNyMTNBlkXysRsLbV1z1O0CILrAqECHH6VoS0TLYKBb65Gc18WUEi06SzHv
# VYh/bfCn6Et35McjQht0F4SbMNSHKTIX+DjMPBcCKzaVKlmBLiYg1P5bUiAcTG4z
# zLY/Y97zZbbWPjKTlN4qT5q/kKy/m8KBnuvuf9F8VbC2RqN+zA==
# SIG # End signature block
