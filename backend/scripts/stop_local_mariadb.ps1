# Arquivo: backend/scripts/stop_local_mariadb.ps1
# Descricao: finaliza processos locais do MariaDB usados no desenvolvimento.

$ErrorActionPreference = "Stop"

$processos = Get-Process -Name mariadbd -ErrorAction SilentlyContinue

if (-not $processos) {
  Write-Host "Nenhum processo mariadbd em execucao foi encontrado."
  exit 0
}

# Notas de manutencao:
# - Objetivo: manter automacoes locais previsiveis.
# - Cuidado: preservar codigos de retorno e mensagens operacionais.
# - Ao alterar: validar caminho de executaveis e parametros.

foreach ($processo in $processos) {
  Stop-Process -Id $processo.Id -Force
  Write-Host "Processo MariaDB finalizado (PID=$($processo.Id))."
}

