# Automação CI + PL (Satake)

Gera a Commercial Invoice e o Packing List (XLSX no layout da Satake) a partir de uma ou mais proformas `.odt`.

## Instalação

```bash
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
```

## Uso

```bash
.venv/bin/python gerar.py "Proforma 13954R1.odt" [outra.odt ...] -p pedido.json
```

- A ordem dos ODTs é a ordem dos blocos "Proforma: X" na invoice.
- A saída padrão vai para `saida/CI+PL<nums> <CLIENTE>_<PAIS>_SAL.xlsx` (ou use `-o arquivo.xlsx`).
- `--json dados.json` grava também o JSON extraído das proformas.
- Só o leitor: `.venv/bin/python leitor_odt.py proforma.odt` imprime o JSON.

O programa **para com erro** quando:
- falta um rótulo ou uma coluna na proforma;
- um item tem `Cte. x Valor Unitario ≠ Valor TOTAL`, a soma não bate com o Total de Items ou com o Total EXW, ou existe desconto, frete ou seguro;
- as proformas têm cliente (RUC), dados bancários ou condição de pagamento diferentes. No caso da condição de pagamento, dá para resolver pelo `pedido.json`.

Ao final, o programa reabre o XLSX e confere se `I = H*C`, se o `=SUM` cobre todos os itens, se a soma bate com o Total EXW e se a PL aponta para as linhas certas da CI. Também lista o que ficou como `[COMPLETAR]`.

## Arquivos

| Arquivo | Papel |
|---|---|
| `leitor_odt.py` | ODT → dados/JSON (rótulos e colunas fixas, sem IA) |
| `config/mapeamento.json` | campo → aba/célula, colunas dos itens e layout do template |
| `escritor_xlsx.py` | preenche o template e recria a área de itens, deslocando total e rodapé |
| `validador.py` | confere a proforma antes de escrever e o XLSX depois |
| `gerar.py` | CLI que junta tudo |
| `config/padroes.json` | dados fixos da Satake (origem, moeda, marca, CNPJ, prefixo "SAL") |
| `config/clientes.json` | cadastro por RUC: Attn., Fono, e-mail, destino, consignatário e grafia preferida do nome/endereço |
| `config/traducciones_es.json` | código → descrição em espanhol (opcional) |
| `pedido_exemplo.json` | campos manuais da invoice: data, via, destino, pesos, caixas, dados da PL por item |
| `template/satake_template.xlsx` | cópia do XLSX de exemplo usada como template |

## Regras adotadas (pendências do escopo, todas configuráveis)

1. **Várias proformas:** suportado. Cada uma vira um bloco "Proforma: X" com numeração própria, e o nº da invoice fica `SAL13683+13940+13944/2026`.
2. **"13954" ou "13954R1":** o padrão é sem revisão, como no exemplo. Para manter a revisão, use `"manter_revisao": true` no pedido ou em `padroes.json`.
3. **Condição de pagamento:** vem da proforma. Se `"condicion_pago"` estiver preenchido no pedido, o valor do pedido vale, e o programa avisa que houve substituição.
4. **Tradução:** só acontece para os códigos que estão em `traducciones_es.json`. Os demais mantêm a descrição da proforma, e o programa avisa.
5. **Pedido da comex:** a forma de pagamento entrou na PL (`H20:K22`) e o bloco de consignatário saiu em espanhol nos dois documentos ("Consignatario / Notificar:"). O consignatário é igual ao cliente, a menos que `clientes.json` ou o pedido defina outro.
6. **Validação:** feita com a 13954R1 (9 itens, USD 10.446,10). Ainda falta testar com os ODTs reais 13683, 13940 e 13944.

Outras escolhas:
- Na PL, as colunas B, C e F são fórmulas que apontam para a CI (`=CI!B26`), como no exemplo. A coluna K recebe o NCM de cada item.
- Pesos, dimensões e m³ por item vão em `itens_pl` no pedido. Se os totais não forem informados, o programa soma os pesos dos itens.
- O CNPJ em `CI!B23` estava `03.4183.185/0001-20` no exemplo e foi corrigido para `03.483.185/0001-20`, como aparece no rodapé. A aba de rascunho `Planilha1` é removida.
- A assinatura e os carimbos DRAFT continuam fora da área de impressão, como no template. A assinatura acompanha o rodapé quando ele se desloca.

## Testes

```bash
.venv/bin/python -m unittest discover -s testes -v
```
