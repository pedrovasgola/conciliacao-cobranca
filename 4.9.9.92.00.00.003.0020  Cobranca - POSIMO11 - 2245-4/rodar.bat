@echo off
chcp 65001 >nul
set PYTHONUTF8=1
set PYTHONUNBUFFERED=1
cd /d "%~dp0"
set "CONCILIACAO_RAIZ=%CD%"
if not exist ".venv\Scripts\python.exe" (
  echo Ambiente ainda nao criado. Na pasta do projeto, rode: uv sync
  pause
  exit /b 1
)
if "%~1"=="" (
  echo Cobranca - informe o periodo.
  echo Pasta: %CD%
  echo Excel em: %CD%\saida
  echo.
  ".venv\Scripts\python.exe" -m conciliacao rodar cobranca
  echo.
  pause
) else (
  ".venv\Scripts\python.exe" -m conciliacao %*
)
