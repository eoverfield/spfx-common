$credentials = Get-Credential
$url = "https://%%%-admin.sharepoint.com"
$siteUrl = "https://%%%.sharepoint.com/sites/apps"

Connect-SPOService -Url $url -Credential $credentials

Set-SPOStorageEntity -Site $siteUrl -Key "customProperty" -Value "custom Value" -Description "a description" -Comments "comments"

Get-SPOStorageEntity -Site $siteUrl -Key "customProperty"
