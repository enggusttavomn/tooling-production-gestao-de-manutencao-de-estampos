$sourceXlsx = "C:\Users\260012599\Box\21 - Process Engineer\01 - Projetos Inovação\2 - Banco de Dados\Novos projetos\FICHAS_DE_MANUTENCAO\FICHA (PADRAO ATUALIZADO).xlsx"
$projectXlsx = "frontend\assets\ficha_padrao.xlsx"
$sheetName = "Ficha Padrão"
$outputHtml = "frontend\index.html"
$pollSeconds = 2

if (-not (Test-Path $sourceXlsx)) {
    Write-Error "Arquivo Excel não encontrado: $sourceXlsx"
    exit 1
}

Write-Host "Monitorando: $sourceXlsx"
Write-Host "Saída HTML: $outputHtml"

$lastWrite = (Get-Item $sourceXlsx).LastWriteTimeUtc

while ($true) {
    Start-Sleep -Seconds $pollSeconds
    $currentWrite = (Get-Item $sourceXlsx).LastWriteTimeUtc
    if ($currentWrite -gt $lastWrite) {
        $lastWrite = $currentWrite
        Copy-Item -Path $sourceXlsx -Destination $projectXlsx -Force
        python tools\excel_to_html.py --input $projectXlsx --sheet $sheetName --output $outputHtml
        Write-Host ("Atualizado em: {0}" -f (Get-Date))
    }
}
