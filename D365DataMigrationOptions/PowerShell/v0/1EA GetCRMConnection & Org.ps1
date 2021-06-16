Get-CrmConnection -InteractiveMode
$CRMConn = Get-CrmConnection -InteractiveMode


Get-CrmConnection -OnLineType Office365 -OrganizationName ZRCRMDev1
$CRMConn = Get-CrmConnection –ServerUrl http://<CRM_Server_Host> -Credential $Cred -OrganizationName <OrgName>
$CRMConn = Get-CrmConnection -Credential $Cred -DeploymentRegion NorthAmerica –OnlineType Office365 –OrganizationName <OrgName>

Get-CrmOrganizations -OnLineType Office365

$CRMOrgs = Get-CrmOrganizations –ServerUrl http://<CRM_Server_Host> –Credential $Cred
$CRMOrgs = Get-CrmOrganizations -Credential $Cred -DeploymentRegion NorthAmerica –OnlineType Office365    
$CRMOrgs = Get-CrmOrganizations -Credential $Cred -DeploymentRegion NorthAmerica –OnlineType Office365
    
