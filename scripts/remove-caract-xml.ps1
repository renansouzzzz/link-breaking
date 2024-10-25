param(
    [string]$pageName
)

$inputXmlFile = ""
$outputNewXmlFile = ""

if (-Not (Test-Path $inputXmlFile)) {
    Write-Host "O arquivo XML não foi encontrado em: $inputXmlFile"
    exit
}

try {
    $xmlContent = Get-Content $inputXmlFile -Raw
    Write-Host "Arquivo XML carregado com sucesso."
} catch {
    Write-Host "Erro ao carregar o arquivo XML: $_"
    exit
}

$cleanedContent = [System.Web.HttpUtility]::HtmlDecode($xmlContent)

$cleanedContent = $cleanedContent -replace '&quot;', '"' `
                                     -replace '&amp;#58;', ':' `
                                     -replace '&amp;#160;', ' ' `
                                     -replace '&lt;', '<' `
                                     -replace '&gt;', '>' `
                                     -replace '&amp;', '&'

try {
    Set-Content -Path $outputNewXmlFile -Value $cleanedContent -Encoding UTF8
    Write-Host "Conteúdo limpo salvo em: $outputNewXmlFile"
    & .\get-links-by-content-xlsx-upd3.ps1 -pageName $pageName
    Remove-Item -Path $inputXmlFile -Force
} catch {
    Write-Host "Erro ao salvar o arquivo limpo: $_"
}
