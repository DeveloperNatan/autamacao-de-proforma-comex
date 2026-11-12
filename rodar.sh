#!/usr/bin/env bash
# Gera a CI + PL com todos os .odt da pasta "proformas".
cd "$(dirname "$0")"
if [ ! -x .venv/bin/python ]; then
  echo "Primeira vez: instalando bibliotecas..."
  python3 -m venv .venv && .venv/bin/pip install -q -r requirements.txt || exit 1
fi
shopt -s nullglob
odts=(proformas/*.odt)
if [ ${#odts[@]} -eq 0 ]; then echo "Coloque os arquivos .odt na pasta 'proformas'."; exit 1; fi
.venv/bin/python gerar.py "${odts[@]}" -p pedido.json
