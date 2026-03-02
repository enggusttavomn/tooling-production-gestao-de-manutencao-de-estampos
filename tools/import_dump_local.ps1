param(
  [Parameter(Mandatory = $true)]
  [string]$DumpPath
)

$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $PSScriptRoot
$clientExe = "C:\Program Files\MariaDB 12.2\bin\mariadb.exe"
$resolvedDump = (Resolve-Path $DumpPath).Path
$tmpDir = Join-Path $projectRoot "_tmp"

if (-not (Test-Path $clientExe)) {
  throw "Cliente MariaDB nao encontrado em: $clientExe"
}

if (-not (Test-Path $resolvedDump)) {
  throw "Dump nao encontrado: $resolvedDump"
}

$baseName = [System.IO.Path]::GetFileNameWithoutExtension($resolvedDump)
if (-not (Test-Path $tmpDir)) {
  New-Item -ItemType Directory -Path $tmpDir | Out-Null
}

$sanitizedDump = Join-Path $tmpDir "${baseName}_nodefiner.sql"
$rawSql = Get-Content -Path $resolvedDump -Raw
$rawSql = $rawSql `
  -replace '/\*![0-9]{5} DEFINER=`[^`]+`@`[^`]+`\s*\*/', '' `
  -replace 'DEFINER=`[^`]+`@`[^`]+`', ''
Set-Content -Path $sanitizedDump -Value $rawSql -Encoding UTF8

$dumpForSql = $sanitizedDump -replace "\\", "/"
$command = "SOURCE `"$dumpForSql`";"

& $clientExe --skip-ssl -h 127.0.0.1 -P 3306 -u dev_local -pdev_local_123 app_ferramental -e $command
if ($LASTEXITCODE -ne 0) {
  throw "Falha ao importar dump no MariaDB local (exit code $LASTEXITCODE)."
}

Write-Host "Import concluido para app_ferramental a partir de: $resolvedDump"
