# Automação de Invoice e Packing List (Comex)

## O problema

Montar a Commercial Invoice (CI) e o Packing List (PL) a partir da proforma é um trabalho manual: copiar cliente, itens, quantidades e preços de um documento para uma planilha. Isso toma tempo e abre espaço para erro de digitação.

## A solução

Um programa em Python lê a proforma (`.odt`) e preenche sozinho uma planilha Excel modelo, com as abas CI e PL. O caminho é este:

1. **Leitura:** o programa extrai os dados da proforma procurando os rótulos fixos do documento ("Nombre del cliente:", "RUC:" etc.) e as colunas da tabela de itens. Não usa IA, então o mesmo arquivo sempre gera o mesmo resultado.
2. **Mapeamento:** um arquivo de configuração (`config/mapeamento.json`) diz em qual aba e em qual célula cada dado deve entrar.
3. **Escrita:** os dados vão para uma cópia da planilha modelo. Os totais ficam como fórmulas do Excel, e o rodapé desce ou sobe conforme a quantidade de itens.
4. **Conferência:** o programa confere se o total da planilha bate com o total da proforma.

Se faltar algum dado ou se os valores não baterem, o programa para e avisa, em vez de gerar uma planilha errada. Ele também junta várias proformas numa invoice só.

## Como adaptar para outro uso

Dá para usar com outros documentos e outras planilhas sem mexer no código principal:

- **Outra planilha modelo:** troque o arquivo em `template/` e ajuste as células em `config/mapeamento.json`.
- **Outra empresa ou outros clientes:** edite `config/padroes.json` e `config/clientes.json`.
- **Outro modelo de proforma:** ajuste os rótulos e os nomes das colunas no início de `leitor_odt.py`.

## Como rodar

Precisa ter o Python 3.13 instalado. No Windows, marque a opção "Add Python to PATH" na instalação.

1. Coloque as proformas na pasta `proformas`.
2. Se quiser, preencha os dados manuais (data, pesos, caixas) no `pedido.json`.
3. Rode `./rodar.sh` no Linux ou dê dois cliques em `rodar.bat` no Windows.
4. A planilha pronta aparece na pasta `saida`.

O manual completo está em `LEIAME.md`.
