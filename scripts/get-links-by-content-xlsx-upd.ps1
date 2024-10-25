param(
    [string]$pageName
)

$xmlFiles = @(
    ""
)
$outputExcelFile = ""

if (-Not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "O módulo ImportExcel não está instalado. Execute 'Install-Module -Name ImportExcel' para instalá-lo."
    exit
}

foreach ($xmlFile in $xmlFiles) {
    try {
        $xmlContent = Get-Content $xmlFile -Raw
        Write-Host "Arquivo XML '$xmlFile' carregado com sucesso."
    } catch {
        Write-Host "Erro ao carregar o arquivo XML: $_"
        continue
    }

    $links = @()
    $pattern = '<a\s+href="([^"]+)"[^>]*>(.*?)<\/a>'
    if ($xmlContent -match $pattern) {
        $matches = [regex]::Matches($xmlContent, $pattern)
        foreach ($match in $matches) {
            $link = $match.Groups[1].Value
            $name = $match.Groups[2].Value -replace '<\/?u>', ''
            $links += [PSCustomObject]@{
                "Título da Página" = $pageName -replace '.aspx', ''
                "Título do Link" = $name
                "URL do Link"    = $link
            }
        }
    }

    $pageName = [System.IO.Path]::GetFileNameWithoutExtension($xmlFile) -replace '.aspx', ''

    if ($links.Count -eq 0) {
        Write-Host "Nenhum link encontrado no conteúdo do arquivo '$xmlFile'."
    } else {
        if (Test-Path $outputExcelFile) {
            $links | Export-Excel -Path $outputExcelFile -WorkSheetname $pageName -AutoSize -Append
        } else {
            $links | Export-Excel -Path $outputExcelFile -WorkSheetname $pageName -AutoSize
        }
        Write-Host "Links extraídos da página '$pageName' e salvos na planilha: $outputExcelFile"
    }

    Remove-Item -Path $xmlFile -Force

}
