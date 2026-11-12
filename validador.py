"""Validações: antes de escrever (proformas) e depois de escrever (XLSX gerado)."""
from __future__ import annotations

from decimal import Decimal

import openpyxl

CENT = Decimal("0.01")


class ErroValidacao(Exception):
    pass


def validar_proforma(pf: dict) -> Decimal:
    """Confere item a item e o total. Retorna o Total EXW."""
    erros = []
    soma = Decimal(0)
    for it in pf["itens"]:
        calc = (it["cantidad"] * it["valor_unitario"]).quantize(CENT)
        if calc != it["valor_total"].quantize(CENT):
            erros.append(
                f"item {it['codigo']}: {it['cantidad']} x {it['valor_unitario']} = {calc}, "
                f"mas a proforma diz {it['valor_total']}"
            )
        soma += calc
    t = pf["totais"]
    if soma != t["total_items"].quantize(CENT):
        erros.append(f"soma dos itens {soma} difere do 'Total de Items' {t['total_items']}")
    exw = (t["total_items"] + t["flete"] + t["seguro"]).quantize(CENT)
    if exw != t["total_exw"].quantize(CENT):
        erros.append(f"Total de Items + Flete + Seguro = {exw}, mas Total EXW = {t['total_exw']}")
    if t["flete"] or t["seguro"]:
        erros.append("proforma tem Flete/Seguro; a CI no layout atual só tem o total EXW dos itens")
    if erros:
        raise ErroValidacao(f"Proforma {pf['numero']} inconsistente:\n  - " + "\n  - ".join(erros))
    return t["total_exw"]


def validar_xlsx(caminho, layout: dict, proformas: list[dict]) -> list[str]:
    """Reabre o XLSX gerado e confere estrutura, fórmulas e totais.

    Retorna a lista de avisos (campos a completar). Levanta ErroValidacao
    se algo estiver errado de fato.
    """
    wb = openpyxl.load_workbook(caminho)
    ci, pl = wb[layout["CI"]["aba"]], wb[layout["PL"]["aba"]]
    erros: list[str] = []

    esperado_total = sum((pf["totais"]["total_exw"] for pf in proformas), Decimal(0))
    n_itens = sum(len(pf["itens"]) for pf in proformas)

    # --- CI: linhas de item ------------------------------------------------
    ini = layout["CI"]["inicio_itens"]
    linhas_item, soma, r = [], Decimal(0), ini
    while True:
        a = ci.cell(r, 1).value
        if isinstance(a, str) and a.startswith("Proforma:"):
            r += 1
            continue
        if isinstance(a, int):
            linhas_item.append(r)
            c, h, i = ci[f"C{r}"].value, ci[f"H{r}"].value, ci[f"I{r}"].value
            if i != f"=H{r}*C{r}":
                erros.append(f"CI!I{r} deveria ser =H{r}*C{r}, está {i!r}")
            soma += (Decimal(str(c)) * Decimal(str(h))).quantize(CENT)
            r += 1
            continue
        break
    linha_total = r
    if len(linhas_item) != n_itens:
        erros.append(f"CI tem {len(linhas_item)} itens, proformas têm {n_itens}")
    if soma != esperado_total.quantize(CENT):
        erros.append(f"Soma calculada da CI {soma} difere do Total EXW das proformas {esperado_total}")
    f_total = ci[f"I{linha_total}"].value
    esperado_f = f"=SUM(I{linhas_item[0]}:I{linhas_item[-1]})" if linhas_item else None
    if f_total != esperado_f:
        erros.append(f"CI!I{linha_total} deveria ser {esperado_f}, está {f_total!r}")

    # --- PL: referências à CI ---------------------------------------------
    ini_pl = layout["PL"]["inicio_itens"]
    linhas_pl = [
        rr for rr in range(ini_pl, ini_pl + n_itens + len(proformas) + 1)
        if isinstance(pl[f"B{rr}"].value, str) and pl[f"B{rr}"].value.startswith("=CI!B")
    ]
    if len(linhas_pl) != n_itens:
        erros.append(f"PL tem {len(linhas_pl)} itens, CI tem {n_itens}")
    for rr_pl, rr_ci in zip(linhas_pl, linhas_item):
        for col_pl, col_ci in (("B", "B"), ("C", "E"), ("F", "C")):
            v = pl[f"{col_pl}{rr_pl}"].value
            if v != f"=CI!{col_ci}{rr_ci}":
                erros.append(f"PL!{col_pl}{rr_pl} deveria apontar para CI!{col_ci}{rr_ci}, está {v!r}")
        if not pl[f"K{rr_pl}"].value:
            erros.append(f"PL!K{rr_pl} sem NCM")

    if erros:
        raise ErroValidacao("XLSX gerado com problemas:\n  - " + "\n  - ".join(erros))

    # --- Avisos: campos manuais pendentes ---------------------------------
    avisos = []
    marcador = layout.get("marcador_pendente", "[COMPLETAR]")
    for ws in (ci, pl):
        for row in ws.iter_rows():
            for cel in row:
                if cel.value is not None and marcador in str(cel.value):
                    avisos.append(f"{ws.title}!{cel.coordinate}: {cel.value}")
    return avisos
