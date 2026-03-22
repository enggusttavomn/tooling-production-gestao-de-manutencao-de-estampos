# Arquivo: backend/scripts/watch_excel.ps1
# Descricao: monitora a planilha fonte e atualiza o HTML de visualizacao local.

$raizProjeto = (Resolve-Path (Join-Path $PSScriptRoot "..\..")).Path
$exePython = Join-Path $raizProjeto ".venv\Scripts\python.exe"
$origemXlsx = "C:\Users\260012599\Box\21 - Process Engineer\01 - Projetos Inovacao\2 - Banco de Dados\Novos projetos\FICHAS_DE_MANUTENCAO\FICHA (PADRAO ATUALIZADO).xlsx"
$destinoXlsx = Join-Path $raizProjeto "backend\data\ficha_padrao.xlsx"
$nomeAba = "Ficha Padrao"
$saidaHtml = Join-Path $raizProjeto "frontend\index.html"
$intervaloSegundos = 2

if (-not (Test-Path $origemXlsx)) {
# Notas de manutencao:
# - Objetivo: manter automacoes locais previsiveis.
# - Cuidado: preservar codigos de retorno e mensagens operacionais.
# - Ao alterar: validar caminho de executaveis e parametros.

    Write-Error "Arquivo Excel nao encontrado: $origemXlsx"
    exit 1
}

Write-Host "Monitorando: $origemXlsx"
Write-Host "Saida HTML: $saidaHtml"

$ultimaEscrita = (Get-Item $origemXlsx).LastWriteTimeUtc

while ($true) {
    Start-Sleep -Seconds $intervaloSegundos
    $escritaAtual = (Get-Item $origemXlsx).LastWriteTimeUtc
    if ($escritaAtual -gt $ultimaEscrita) {
        $ultimaEscrita = $escritaAtual
        Copy-Item -Path $origemXlsx -Destination $destinoXlsx -Force
        if (-not (Test-Path $exePython)) {
            Write-Error "Python da virtualenv nao encontrado: $exePython"
            exit 1
        }
        & $exePython (Join-Path $raizProjeto "backend\utils\excel_to_html.py") --input $destinoXlsx --sheet $nomeAba --output $saidaHtml
        Write-Host ("Atualizado em: {0}" -f (Get-Date))
    }
}

