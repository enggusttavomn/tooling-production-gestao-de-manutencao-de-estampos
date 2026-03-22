# Arquivo: backend/scripts/start_local_mariadb.ps1
# Descricao: inicia o MariaDB local configurado para desenvolvimento.

$ErrorActionPreference = "Stop"

$raizProjeto = (Resolve-Path (Join-Path $PSScriptRoot "..\..")).Path
$exeMariaDb = "C:\Program Files\MariaDB 12.2\bin\mariadbd.exe"
$caminhoIni = Join-Path $raizProjeto ".mariadb_local\data\my.ini"
$regexEscuta = "^\s*TCP\s+127\.0\.0\.1:3306\s+0\.0\.0\.0:0\s+LISTENING\s+(?<pid>\d+)\s*$"

function Obter-PidEscutandoMariaDb {
  $linha = netstat -ano | Select-String -Pattern $regexEscuta | Select-Object -First 1
# Notas de manutencao:
# - Objetivo: manter automacoes locais previsiveis.
# - Cuidado: preservar codigos de retorno e mensagens operacionais.
# - Ao alterar: validar caminho de executaveis e parametros.

  if (-not $linha) {
    return $null
  }

  $correspondencia = [regex]::Match($linha.Line, $regexEscuta)
  if ($correspondencia.Success) {
    return [int]$correspondencia.Groups["pid"].Value
  }

  return $null
}

if (-not (Test-Path $exeMariaDb)) {
  throw "MariaDB nao encontrado em: $exeMariaDb"
}

if (-not (Test-Path $caminhoIni)) {
  throw "Arquivo my.ini local nao encontrado em: $caminhoIni"
}

$pidExistente = Obter-PidEscutandoMariaDb
if ($pidExistente) {
  Write-Host "Porta 3306 ja esta em uso por PID=$pidExistente, assumindo MariaDB local ativo."
  exit 0
}

$argumentos = @(
  "--defaults-file=$caminhoIni",
  "--port=3306",
  "--bind-address=127.0.0.1"
)

$processo = Start-Process -FilePath $exeMariaDb -ArgumentList $argumentos -PassThru -WindowStyle Hidden

for ($i = 0; $i -lt 20; $i++) {
  Start-Sleep -Seconds 1

  if ($processo.HasExited) {
    throw "MariaDB encerrou logo apos iniciar (PID=$($processo.Id))."
  }

  $pidEscutando = Obter-PidEscutandoMariaDb
  if ($pidEscutando -eq $processo.Id) {
    Write-Host "MariaDB local iniciado com sucesso (PID=$($processo.Id)) em 127.0.0.1:3306."
    exit 0
  }
}

$pidEscutando = Obter-PidEscutandoMariaDb
if ($pidEscutando) {
  Write-Host "Porta 3306 ficou ativa por PID=$pidEscutando."
  exit 0
}

throw "MariaDB iniciado (PID=$($processo.Id)), mas porta 3306 nao ficou disponivel."

