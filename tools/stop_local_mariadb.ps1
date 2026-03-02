$ErrorActionPreference = "Stop"

$procs = Get-Process -Name mariadbd -ErrorAction SilentlyContinue

if (-not $procs) {
  Write-Host "Nenhum processo mariadbd em execucao foi encontrado."
  exit 0
}

foreach ($p in $procs) {
  Stop-Process -Id $p.Id -Force
  Write-Host "Processo MariaDB finalizado (PID=$($p.Id))."
}
