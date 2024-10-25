param(
    [string]$pageName
)

$siteUrl = ""
$outputFile = ""

$appId = ""
$appSecret = ""

Connect-PnPOnline -Url $siteUrl -UseWebLogin

$page = Get-PnPFile -Url "" -AsString

$xmlDoc = New-Object System.Xml.XmlDocument
$root = $xmlDoc.CreateElement("SharepointPage")
$xmlDoc.AppendChild($root)

$pageContent = $xmlDoc.CreateElement("PageContent")
$pageContent.InnerText = $page
$root.AppendChild($pageContent)

$xmlDoc.Save($outputFile)

Write-Host "A página foi exportada com sucesso para $outputFile"

& .\remove-caract-xml2.ps1 -pageName $pageName