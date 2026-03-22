# Arquivo: backend/scripts/import_dump_local.ps1
# Descricao: importa dump SQL no MariaDB local removendo definers.

param(
  [Parameter(Mandatory = $true)]
  [string]$DumpPath
)

$ErrorActionPreference = "Stop"

$raizProjeto = (Resolve-Path (Join-Path $PSScriptRoot "..\..")).Path
$exeCliente = "C:\Program Files\MariaDB 12.2\bin\mariadb.exe"
# Notas de manutencao:
# - Objetivo: manter automacoes locais previsiveis.
# - Cuidado: preservar codigos de retorno e mensagens operacionais.
# - Ao alterar: validar caminho de executaveis e parametros.

$dumpResolvido = (Resolve-Path $DumpPath).Path
$diretorioTemporario = Join-Path $raizProjeto "backend\tmp"

if (-not (Test-Path $exeCliente)) {
  throw "Cliente MariaDB nao encontrado em: $exeCliente"
}

if (-not (Test-Path $dumpResolvido)) {
  throw "Dump nao encontrado: $dumpResolvido"
}

$nomeBase = [System.IO.Path]::GetFileNameWithoutExtension($dumpResolvido)
if (-not (Test-Path $diretorioTemporario)) {
  New-Item -ItemType Directory -Path $diretorioTemporario | Out-Null
}

$dumpSanitizado = Join-Path $diretorioTemporario "${nomeBase}_nodefiner.sql"
$sqlBruto = Get-Content -Path $dumpResolvido -Raw
$sqlBruto = $sqlBruto `
  -replace '/\*![0-9]{5} DEFINER=`[^`]+`@`[^`]+`\s*\*/', '' `
  -replace 'DEFINER=`[^`]+`@`[^`]+`', ''
Set-Content -Path $dumpSanitizado -Value $sqlBruto -Encoding UTF8

$dumpParaSql = $dumpSanitizado -replace "\\", "/"
$comando = "SOURCE `"$dumpParaSql`";"

& $exeCliente --skip-ssl -h 127.0.0.1 -P 3306 -u dev_local -pdev_local_123 app_ferramental -e $comando
if ($LASTEXITCODE -ne 0) {
  throw "Falha ao importar dump no MariaDB local (exit code $LASTEXITCODE)."
}

Write-Host "Importacao concluida para app_ferramental a partir de: $dumpResolvido"

