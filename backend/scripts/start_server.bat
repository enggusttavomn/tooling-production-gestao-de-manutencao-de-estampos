:: Arquivo: backend/scripts/start_server.bat
:: Descricao: inicializa o ambiente local da aplicacao Flask com reinicio automatico.
::
:: Fluxo resumido:
:: 1) Resolve caminhos absolutos do projeto (venv, app, scripts).
:: 2) Garante pasta de logs local.
:: 3) Opcionalmente sobe MariaDB local antes da API.
:: 4) Inicia Flask e registra crash em log para diagnostico.
:: 5) Reinicia automaticamente quando AUTO_RESTART=1.
@echo off
setlocal

:: Notas de manutencao:
:: - Objetivo: manter inicializacao local simples e reproduzivel.
:: - Cuidado: preservar variaveis de ambiente e codigos de erro.
:: - Ao alterar: testar ciclo iniciar/parar/reiniciar.

set "SCRIPT_DIR=%~dp0"
for %%I in ("%SCRIPT_DIR%..\..") do set "PROJECT_ROOT=%%~fI"
set "PYTHON_EXE=%PROJECT_ROOT%\.venv\Scripts\python.exe"
set "APP_ENTRY=%PROJECT_ROOT%\backend\api\app.py"
set "DB_START_SCRIPT=%PROJECT_ROOT%\backend\scripts\start_local_mariadb.ps1"
set "LOG_DIR=%PROJECT_ROOT%\backend\logs"
set "CRASH_LOG=%LOG_DIR%\server_crash.log"

if "%FLASK_HOST%"=="" set "FLASK_HOST=0.0.0.0"
if "%PORT%"=="" set "PORT=5000"
if "%FLASK_DEBUG%"=="" set "FLASK_DEBUG=0"
if "%AUTO_RESTART%"=="" set "AUTO_RESTART=1"
if "%RESTART_DELAY_SEC%"=="" set "RESTART_DELAY_SEC=3"

echo Iniciando servidor...
echo Usando: "%PYTHON_EXE%"

:: Logs de runtime ficam no backend para manter raiz limpa.
if not exist "%LOG_DIR%" mkdir "%LOG_DIR%"

if not exist "%PYTHON_EXE%" (
  echo.
  echo ERRO: venv do Python nao encontrada.
  echo Esperado em: "%PYTHON_EXE%"
  echo.
  echo Dica: crie a venv ou atualize este script.
  pause
  exit /b 1
)

if not exist "%APP_ENTRY%" (
  echo.
  echo ERRO: app.py nao encontrado.
  echo Esperado em: "%APP_ENTRY%"
  pause
  exit /b 1
)

:: Se o script de MariaDB existir, tenta subir o banco local antes da API.
if exist "%DB_START_SCRIPT%" (
  echo Iniciando MariaDB local...
  powershell -ExecutionPolicy Bypass -File "%DB_START_SCRIPT%"
  if errorlevel 1 (
    echo.
    echo ERRO: falha ao iniciar o MariaDB local.
    pause
    exit /b 1
  )
) else (
  echo AVISO: script de inicio do MariaDB nao encontrado em "%DB_START_SCRIPT%".
)

:RUN_SERVER
cd /d "%PROJECT_ROOT%"
echo.
echo [%date% %time%] Iniciando Flask em %FLASK_HOST%:%PORT% (debug=%FLASK_DEBUG%)...
"%PYTHON_EXE%" "%APP_ENTRY%"
set "EXIT_CODE=%ERRORLEVEL%"

if "%EXIT_CODE%"=="0" goto SERVER_STOPPED
echo [%date% %time%] Flask exited with code %EXIT_CODE%.>>"%CRASH_LOG%"

if not "%AUTO_RESTART%"=="1" goto SERVER_STOPPED
echo [%date% %time%] Falha detectada (codigo %EXIT_CODE%). Reiniciando em %RESTART_DELAY_SEC%s...
timeout /t %RESTART_DELAY_SEC% /nobreak >nul
goto RUN_SERVER

:SERVER_STOPPED
echo.
echo Servidor parado ou falha ao iniciar.
pause