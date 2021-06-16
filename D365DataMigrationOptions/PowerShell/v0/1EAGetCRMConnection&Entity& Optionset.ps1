Param(

    $ConnectionString = "AuthType=ClientSecret;url=https://ztdev.crm11.dynamics.com/;ClientId=c2b575f9-ef65-490b-8741-800f067f2111;ClientSecret=22-OC1YovwX38F~13.K-I1gyusuh4_iUZo",

    $EntityLogicalName = "hsl_configsetting",

    $OptionSetAttributeName = "statuscode"
)

# Import the Dynamics XRM PowerShell module
# https://www.powershellgallery.com/packages/Microsoft.Xrm.Tooling.CrmConnector.PowerShell/3.3.0.887
Import-Module Microsoft.Xrm.Tooling.CrmConnector.PowerShell;
# Get a CrmServiceClient to communicate with Dynamics CRM
$CrmClient = Get-CrmConnection -ConnectionString $ConnectionString;

# Create a RetrieveAttributeRequest to fetch Attribute metadata
$AttributeRequest = [Microsoft.Xrm.Sdk.Messages.RetrieveAttributeRequest]::new();
$AttributeRequest.EntityLogicalName = $EntityLogicalName;
$AttributeRequest.LogicalName = $OptionSetAttributeName;
$AttributeRequest.RetrieveAsIfPublished = $True;
$AttributeResponse = [Microsoft.Xrm.Sdk.Messages.RetrieveAttributeResponse]$CrmClient.Execute($AttributeRequest);

# Get the Value/Label pairs and print them to the console
$AttributeResponse.AttributeMetadata.OptionSet.Options `
    | Select-Object -Property `
        @{Name = "Value"; Expression={$_.Value}},`
        @{Name = "Label"; Expression={$_.Label.UserLocalizedLabel.Label}};

# Close the connection to CRM
$CrmClient.Dispose();