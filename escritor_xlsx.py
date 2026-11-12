"""Escreve CI e PL no template da Satake a partir do contexto + mapeamento.

Passos:
 1. grava os campos de cabeçalho/rodapé nas coordenadas do template;
 2. reconstrói a área de itens de cada aba (blocos "Proforma: X" + itens),
    copiando o estilo das linhas-modelo e deslocando total e rodapé;
 3. totais sempre como fórmula (=H*C e =SUM); nada de valor copiado.
"""
from __future__ import annotations

import re
from copy import copy
from decimal import Decimal
from pathlib import Path

import openpyxl
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.styles import Alignment
from openpyxl.worksheet.pagebreak import RowBreak

PLACEHOLDER = re.compile(r"\{([a-zA-Z0-9_.]+)\}")


class ErroEscrita(Exception):
    pass


# --------------------------------------------------------------------------
# Resolução de {chaves}
# --------------------------------------------------------------------------
def _pega(ctx: dict, chave: str):
    atual = ctx
    for parte in chave.split("."):
        if isinstance(atual, dict) and parte in atual:
            atual = atual[parte]
        else:
            return None
    return atual


def _num(v):
    if isinstance(v, Decimal):
        return int(v) if v == v.to_integral_value() else float(v)
    return v


def resolver(modelo, ctx: dict, pendente: str | None):
    """Substitui {chaves}. Se o modelo é só '{x}', devolve o valor com o tipo
    original (número continua número). Ausente -> `pendente` (ou None)."""
    if not isinstance(modelo, str):
        return modelo
    m = PLACEHOLDER.fullmatch(modelo)
    if m:
        v = _pega(ctx, m.group(1))
        return pendente if v in (None, "") else _num(v)
    faltou = False

    def troca(mm):
        nonlocal faltou
        v = _pega(ctx, mm.group(1))
        if v in (None, ""):
            faltou = True
            return pendente or ""
        return str(_num(v))

    texto = PLACEHOLDER.sub(troca, modelo)
    if faltou and pendente is None:
        return None
    return texto


# --------------------------------------------------------------------------
# Campos fixos
# --------------------------------------------------------------------------
def _escreve_campo(ws, campo: dict, ctx: dict, marcador: str):
    cel = ws[campo["celula"]]
    pend = None if campo.get("opcional") else marcador
    valor = resolver(campo["valor"], ctx, pend)
    if campo.get("rotulo"):
        f = cel.font
        normal = InlineFont(rFont=f.name, sz=f.sz, b=False, color=f.color)
        cel.value = CellRichText(campo["rotulo"], TextBlock(normal, str(valor or "")))
        cel.font = copy(f) if f.b else f.copy(b=True)
    else:
        cel.value = valor
    if campo.get("mesclar"):
        ws.merge_cells(campo["mesclar"])
    if campo.get("quebrar_texto"):
        cel.alignment = Alignment(wrap_text=True, vertical="top", horizontal="left")


# --------------------------------------------------------------------------
# Área de itens
# --------------------------------------------------------------------------
def _foto_linha(ws, r: int, max_col: int) -> dict:
    """Valores, estilos, altura e mesclagens de uma linha."""
    return {
        "celulas": {
            c: (copy(ws.cell(r, c).value), copy(ws.cell(r, c)._style))
            for c in range(1, max_col + 1)
            if (r, c) in ws._cells
        },
        "altura": ws.row_dimensions[r].height if r in ws.row_dimensions else None,
        "mescla": [
            (m.min_col, m.max_col, m.max_row - m.min_row)
            for m in ws.merged_cells.ranges
            if m.min_row == r
        ],
    }


def _aplica_foto(ws, foto: dict, r: int, com_valores: bool):
    for c, (v, st) in foto["celulas"].items():
        cel = ws.cell(r, c)
        cel._style = copy(st)
        cel.value = copy(v) if com_valores else None
    if foto["altura"]:
        ws.row_dimensions[r].height = foto["altura"]
    for c1, c2, dr in foto["mescla"]:
        ws.merge_cells(start_row=r, start_column=c1, end_row=r + dr, end_column=c2)


def _reconstroi(ws, lay: dict, blocos: list[dict], ctx: dict, mapa_itens: dict,
                mapa_grupo: dict, linhas_ci: list[int] | None) -> list[int]:
    ini, total_old = lay["inicio_itens"], lay["linha_total"]
    max_col = max(ws.max_column, 12)
    max_row = ws.max_row

    foto_grupo = _foto_linha(ws, lay["modelo_grupo"], max_col)
    foto_item = _foto_linha(ws, lay["modelo_item"], max_col)
    # A linha-modelo de item pode estar dentro de mesclagens verticais (PL
    # agrupa caixas). Para item avulso só vale mesclagem horizontal.
    foto_item["mescla"] = [(a, b, 0) for a, b, _ in foto_item["mescla"]]
    foto_total = _foto_linha(ws, total_old, max_col)
    rodape = [_foto_linha(ws, r, max_col) for r in range(total_old + 1, max_row + 1)]

    # limpa da área de itens até o fim
    for m in list(ws.merged_cells.ranges):
        if m.max_row >= ini:
            if m.min_row < ini:
                raise ErroEscrita(f"{ws.title}: mesclagem {m} cruza o início da área de itens")
            ws.unmerge_cells(str(m))
    for chave in [k for k in ws._cells if k[0] >= ini]:
        del ws._cells[chave]
    for r in [r for r in ws.row_dimensions if r >= ini]:
        del ws.row_dimensions[r]

    r = ini
    linhas_item: list[int] = []
    n = 0
    for bloco in blocos:
        _aplica_foto(ws, foto_grupo, r, com_valores=False)
        for col, modelo in mapa_grupo.items():
            ws[f"{col}{r}"] = resolver(modelo, {**ctx, "proforma": bloco}, None)
        r += 1
        for seq, item in enumerate(bloco["itens"], start=1):
            _aplica_foto(ws, foto_item, r, com_valores=False)
            local = {**ctx, "item": {**item, "seq": seq}, "proforma": bloco,
                     "linha": r, "linha_ci": linhas_ci[n] if linhas_ci else None}
            for col, modelo in mapa_itens.items():
                modelo = modelo.replace("{linha}", str(r))
                if local["linha_ci"]:
                    modelo = modelo.replace("{linha_ci}", str(local["linha_ci"]))
                ws[f"{col}{r}"] = resolver(modelo, local, None)
            linhas_item.append(r)
            n += 1
            r += 1

    total_new = r
    _aplica_foto(ws, foto_total, total_new, com_valores=True)
    for col, func in lay.get("formulas_total", {}).items():
        ws[f"{col}{total_new}"] = f"={func}({col}{linhas_item[0]}:{col}{linhas_item[-1]})"

    for i, foto in enumerate(rodape, start=1):
        _aplica_foto(ws, foto, total_new + i, com_valores=True)

    delta = total_new - total_old
    # imagens do rodapé (assinatura) acompanham o deslocamento
    for img in ws._images:
        a = img.anchor
        if hasattr(a, "_from") and a._from.row >= total_old:
            a._from.row += delta
            if getattr(a, "to", None) is not None:
                a.to.row += delta

    c1, c2 = lay["colunas_impressao"].split(":")
    ws.print_area = f"{c1}1:{c2}{lay['ultima_linha_impressao'] + delta}"
    ws.row_breaks = RowBreak()
    return linhas_item


# --------------------------------------------------------------------------
# API
# --------------------------------------------------------------------------
def escrever(mapeamento: dict, ctx: dict, blocos: list[dict], destino, base_dir) -> None:
    wb = openpyxl.load_workbook(base_dir / mapeamento["template"], rich_text=True)
    marcador = mapeamento["marcador_pendente"]
    lay = mapeamento["layout"]

    for aba in mapeamento.get("remover_abas", []):
        if aba in wb.sheetnames:
            del wb[aba]
    for nome in ("CI", "PL"):
        if lay[nome]["aba"] not in wb.sheetnames:
            raise ErroEscrita(f"Aba '{lay[nome]['aba']}' não existe no template")

    # Estilo de células novas (ex.: PL!H20) herdado da célula de cima.
    for campo in mapeamento["campos"]:
        ws = wb[lay[campo["aba"]]["aba"]]
        cel = ws[campo["celula"]]
        if (cel.row, cel.column) not in ws._cells or not cel.has_style:
            acima = ws.cell(cel.row - 1, cel.column)
            cel._style = copy(acima._style)
        _escreve_campo(ws, campo, ctx, marcador)

    linhas_ci = _reconstroi(wb[lay["CI"]["aba"]], lay["CI"], blocos, ctx,
                            mapeamento["itens"]["CI"], mapeamento["grupo"], None)
    _reconstroi(wb[lay["PL"]["aba"]], lay["PL"], blocos, ctx,
                mapeamento["itens"]["PL"], mapeamento["grupo"], linhas_ci)

    wb.calculation.fullCalcOnLoad = True  # Excel/LibreOffice recalculam ao abrir
    wb.save(destino)
    _corrige_quebras_cabecalho(destino)


_TAG_HF = re.compile(r"(<(odd|even|first)(Header|Footer)>)(.*?)(</\2\3>)", re.S)


def _corrige_quebras_cabecalho(caminho) -> None:
    """O openpyxl grava as quebras de linha do cabeçalho/rodapé de página como
    o texto literal '_x000a_'. Troca por quebra real só dentro dessas tags."""
    import shutil
    import tempfile
    import zipfile

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx", dir=str(Path(caminho).parent))
    tmp.close()
    with zipfile.ZipFile(caminho) as zin, zipfile.ZipFile(tmp.name, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            dados = zin.read(info.filename)
            if info.filename.startswith("xl/worksheets/sheet") and b"_x000" in dados:
                xml = dados.decode("utf-8")
                xml = _TAG_HF.sub(
                    lambda m: m.group(1)
                    + m.group(4).replace("_x000d_", "").replace("_x000a_", "&#10;")
                    + m.group(5),
                    xml,
                )
                dados = xml.encode("utf-8")
            zout.writestr(info, dados)
    shutil.move(tmp.name, caminho)
