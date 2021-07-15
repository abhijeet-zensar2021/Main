param (
    $UserEmail = $UserEmail,
    $UserName = $UserName,
    $FeatureBranch = $FeatureBranch    
)

Write-Host "Configure user name and email"

git config --global user.email "$($UserEmail)"
git config --global user.name "$($UserName)"

Write-Host "Checkout Feature  Branch"

git checkout $($FeatureBranch)

Write-Host "Stage all changes"
git add -A

Write-Host "Commit all changes"
git commit -m "Changes committed"

git push origin HEAD:"$($FeatureBranch)"

Write-Host "Package Uploaded to ADO"