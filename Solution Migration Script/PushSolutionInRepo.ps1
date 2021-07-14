Write-Host "Configure user name and email"

git config user.email $(UserEmail)
git config user.name $(UserName)

Write-Host "Checkout Feature  Branch"

git checkout $(FeatureBranch)

Write-Host "Stage all changes"
git add -A

Write-Host "Commit all changes"
git commit -m "Adding file"
git -c http.extraheader="AUTHORIZATION: bearer $(System.AccessToken)" push origin $(FeatureBranch)

Write-Host "Package Uploaded to ADO"