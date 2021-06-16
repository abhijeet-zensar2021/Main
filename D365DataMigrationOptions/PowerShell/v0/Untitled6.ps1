
param(
    #$CrmConnectionString = "AuthType=ClientSecret;url=https://ztdev.crm11.dynamics.com/;ClientId=c2b575f9-ef65-490b-8741-800f067f2111;ClientSecret=22-OC1YovwX38F~13.K-I1gyusuh4_iUZo"
    $CrmConnectionString = "AuthType=Office365;url=https://ztdev.crm11.dynamics.com/;UserName=Elijah@ztd365.onmicrosoft.com;Password=Zensar@1234"
)
#Get-CrmConnection -ConnectionString "AuthType=AD;Url=https://myserver/Contoso;Domain=contosodom;UserName=user1;Password=password"

Get-CrmConnection -ConnectionString "$CrmConnectionString"
#Get-CrmConnection -ConnectionString "AuthType=ClientSecret;url=https://ztdev.crm11.dynamics.com/;ClientId=c2b575f9-ef65-490b-8741-800f067f2111;ClientSecret=22-OC1YovwX38F~13.K-I1gyusuh4_iUZo"
#Get-CrmConnection -ConnectionString "AuthType=Office365;url=https://ztdev.crm11.dynamics.com/;UserName=Elijah@ztd365.onmicrosoft.com;Password=Zensar@1234"