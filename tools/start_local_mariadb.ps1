$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $PSScriptRoot
$mariadbExe = "C:\Program Files\MariaDB 12.2\bin\mariadbd.exe"
$iniPath = Join-Path $projectRoot ".mariadb_local\data\my.ini"

if (-not (Test-Path $mariadbExe)) {
  throw "MariaDB nao encontrado em: $mariadbExe"
}

if (-not (Test-Path $iniPath)) {
  throw "Arquivo my.ini local nao encontrado em: $iniPath"
}

$portInUse = netstat -ano | Select-String "127.0.0.1:3306"
if ($portInUse) {
  Write-Host "Porta 3306 ja esta em uso, assumindo MariaDB local ativo."
  exit 0
}

$args = @(
  "--defaults-file=$iniPath",
  "--port=3306",
  "--bind-address=127.0.0.1"
)

$proc = Start-Process -FilePath $mariadbExe -ArgumentList $args -PassThru -WindowStyle Hidden
Start-Sleep -Seconds 2

$listening = netstat -ano | Select-String ":3306"
if (-not $listening) {
  throw "MariaDB iniciado (PID=$($proc.Id)), mas porta 3306 nao ficou disponivel."
}

Write-Host "MariaDB local iniciado com sucesso (PID=$($proc.Id)) em 127.0.0.1:3306."
