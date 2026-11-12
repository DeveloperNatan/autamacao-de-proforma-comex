@echo off
rem Gera a CI + PL com todos os .odt da pasta "proformas".
cd /d "%~dp0"
setlocal enabledelayedexpansion
if not exist .venv\Scripts\python.exe (
  echo Primeira vez: instalando bibliotecas...
  python -m venv .venv || goto erro
  .venv\Scripts\pip install -q -r requirements.txt || goto erro
)
set ARQS=
for %%f in (proformas\*.odt) do set ARQS=!ARQS! "%%f"
if "!ARQS!"=="" (
  echo Coloque os arquivos .odt na pasta "proformas".
  goto fim
)
.venv\Scripts\python gerar.py !ARQS! -p pedido.json
goto fim
:erro
echo ERRO na instalacao. Verifique se o Python esta instalado.
:fim
pause
