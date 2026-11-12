# Automação de Invoice e Packing List (Comex)

Gera a Commercial Invoice (CI) e o Packing List (PL) em Excel, no layout da Satake, a partir das proformas em `.odt`.

## Como usar

1. Coloque as proformas (`.odt`) na pasta `proformas`.
2. Se quiser, preencha a data, os pesos e as caixas no `pedido.json`.
3. Rode o atalho:
   - Linux: `./rodar.sh`
   - Windows: dois cliques em `rodar.bat`
4. A planilha pronta aparece na pasta `saida`.

## O que precisa ter instalado

- Python 3.13. No Windows, marque a opção "Add Python to PATH" na instalação.

O manual completo está em `LEIAME.md`.
