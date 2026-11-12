"""Leitor determinístico da proforma (.odt) gerada pelo sistema.

Abre o ODT como zip, lê o content.xml e extrai os campos ancorados nos
rótulos fixos e nas colunas fixas da tabela de itens. Se algum rótulo ou
coluna não for encontrado, levanta ErroLeitura (falha ruidosa).

Uso direto:  python leitor_odt.py "Proforma 13954R1.odt"   -> imprime JSON
"""
from __future__ import annotations

import json
import re
import sys
import zipfile
from decimal import Decimal, InvalidOperation
from pathlib import Path

from lxml import etree

NS = {
    "text": "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
    "table": "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
    "office": "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
}
T = "{%s}" % NS["text"]
TB = "{%s}" % NS["table"]


class ErroLeitura(Exception):
    """Rótulo, coluna ou valor esperado não encontrado na proforma."""


# --------------------------------------------------------------------------
# Texto
# --------------------------------------------------------------------------
def _texto(el) -> str:
    """Texto de um elemento ODF respeitando <text:s>, <text:tab> e quebras."""
    partes: list[str] = []

    def walk(e):
        if e.text:
            partes.append(e.text)
        for filho in e:
            tag = filho.tag
            if tag == T + "s":
                partes.append(" " * int(filho.get(T + "c", "1")))
            elif tag == T + "tab":
                partes.append("\t")
            elif tag == T + "line-break":
                partes.append("\n")
            else:
                walk(filho)
            if filho.tail:
                partes.append(filho.tail)

    walk(el)
    return "".join(partes)


def _limpa(s: str) -> str:
    return re.sub(r"[  ]+", " ", s).strip()


def numero_br(s: str, campo: str = "") -> Decimal:
    """'USD1.576,40' -> Decimal('1576.40');  '2,00' -> Decimal('2.00')."""
    bruto = re.sub(r"[^\d,.\-]", "", s or "")
    bruto = bruto.replace(".", "").replace(",", ".")
    try:
        return Decimal(bruto)
    except InvalidOperation:
        raise ErroLeitura(f"Valor numérico inválido em {campo or 'campo'}: {s!r}")


# --------------------------------------------------------------------------
# Rótulos (parágrafos fora das tabelas)
# --------------------------------------------------------------------------
ROTULOS = {
    "numero_proforma": "Numero de Pro forma:",
    "fecha_emision": "Fecha de emisión:",
    "validez": "Validez:",
    "cliente_nombre": "Nombre del cliente:",
    "cliente_direccion": "Dirección:",
    "cliente_ciudad": "Ciudad/UF:",
    "cliente_ruc": "RUC/RUT/NIT:",
    "condicion_pago": "Condición de pago:",
    "fecha_entrega": "Fecha de entrega en la fabrica de Satake:",
    "banco_intermediario": "Banco Intermediario:",
    "banco_beneficiario": "Banco Beneficiario:",
    "iban": "IBAN:",
    "beneficiario": "Beneficiario:",
}
OPCIONAIS = {"validez", "fecha_entrega", "banco_intermediario"}


def _paragrafos_fora_de_tabela(corpo) -> list[str]:
    saida = []
    for p in corpo.iter(T + "p", T + "h"):
        if any(a.tag == TB + "table" for a in p.iterancestors()):
            continue
        saida.append(_texto(p))
    return saida


def _busca_rotulo(paragrafos: list[str], rotulo: str) -> str | None:
    # O rótulo precisa estar no início do parágrafo ou logo após um TAB,
    # para "Beneficiario:" não casar dentro de "Banco Beneficiario:".
    padrao = re.compile(r"(?:^|\t)[  ]*" + re.escape(rotulo) + r"([^\t\n]*)")
    for p in paragrafos:
        m = padrao.search(p)
        if m:
            return _limpa(m.group(1))
    return None


# --------------------------------------------------------------------------
# Tabelas
# --------------------------------------------------------------------------
def _linhas_tabela(tabela) -> list[list[str]]:
    linhas = []
    for tr in tabela.iter(TB + "table-row"):
        cels = []
        for tc in tr:
            if tc.tag not in (TB + "table-cell", TB + "covered-table-cell"):
                continue
            rep = int(tc.get(TB + "number-columns-repeated", "1"))
            cels.extend([_limpa(_texto(tc))] * min(rep, 50))
        linhas.append(cels)
    return linhas


COLUNAS_ITENS = {
    "codigo": "Referencia",
    "descripcion": "Descripción",
    "ncm": "NCM",
    "cantidad": "Cte.",
    "valor_unitario": "Valor Unitario",
    "desconto": "Desc %",
    "valor_total": "Valor TOTAL",
}


def _extrai_itens(tabelas) -> list[dict]:
    for tab in tabelas:
        linhas = _linhas_tabela(tab)
        if not linhas:
            continue
        cab = linhas[0]
        if "Referencia" not in cab or "Descripción" not in cab:
            continue
        idx = {}
        for campo, titulo in COLUNAS_ITENS.items():
            if titulo not in cab:
                raise ErroLeitura(f"Coluna '{titulo}' não encontrada na tabela de itens. Cabeçalho: {cab}")
            idx[campo] = cab.index(titulo)
        itens = []
        for n, ln in enumerate(linhas[1:], start=1):
            if not any(ln):
                continue
            if len(ln) < len(cab):
                raise ErroLeitura(f"Linha {n} da tabela de itens incompleta: {ln}")
            g = {c: ln[i] for c, i in idx.items()}
            desc = re.sub(r"\s*-\s*$", "", g["descripcion"]).strip()  # remove " -" final
            item = {
                "codigo": g["codigo"],
                "descripcion": desc,
                "ncm": g["ncm"],
                "cantidad": numero_br(g["cantidad"], f"Cte. do item {g['codigo']}"),
                "valor_unitario": numero_br(g["valor_unitario"], f"Valor Unitario do item {g['codigo']}"),
                "desconto": numero_br(g["desconto"], f"Desc % do item {g['codigo']}"),
                "valor_total": numero_br(g["valor_total"], f"Valor TOTAL do item {g['codigo']}"),
            }
            if not item["codigo"]:
                raise ErroLeitura(f"Item {n} sem Referencia: {ln}")
            if not re.fullmatch(r"\d{8}", item["ncm"]):
                raise ErroLeitura(f"NCM inválido no item {item['codigo']}: {item['ncm']!r}")
            if item["desconto"] != 0:
                raise ErroLeitura(
                    f"Item {item['codigo']} tem desconto de {item['desconto']}%. "
                    "Desconto ainda não é tratado na CI; ajuste manualmente ou implemente a regra."
                )
            itens.append(item)
        if not itens:
            raise ErroLeitura("Tabela de itens encontrada, mas sem nenhum item.")
        return itens
    raise ErroLeitura("Tabela de itens (cabeçalho 'Referencia | Descripción | ...') não encontrada.")


def _extrai_totais(tabelas) -> dict:
    alvo = {
        "total_items": lambda r: r.startswith("Total de Items"),
        "flete": lambda r: r.startswith("Flete"),
        "seguro": lambda r: r.startswith("Seguro"),
        "total_exw": lambda r: r.startswith("Total EXW"),
    }
    achados: dict = {}
    for tab in tabelas:
        for ln in _linhas_tabela(tab):
            if len(ln) < 2:
                continue
            for campo, teste in alvo.items():
                if campo not in achados and teste(ln[0]):
                    achados[campo] = numero_br(ln[-1], ln[0])
                    if campo == "total_exw":
                        achados["incoterm"] = ln[0].rstrip(":")
    faltando = [c for c in alvo if c not in achados]
    if faltando:
        raise ErroLeitura(f"Totais não encontrados na proforma: {', '.join(faltando)}")
    return achados


# --------------------------------------------------------------------------
# API
# --------------------------------------------------------------------------
def ler_proforma(caminho: str | Path) -> dict:
    caminho = Path(caminho)
    try:
        with zipfile.ZipFile(caminho) as z:
            xml = z.read("content.xml")
    except (zipfile.BadZipFile, KeyError) as e:
        raise ErroLeitura(f"{caminho.name}: não é um ODT válido ({e})")

    raiz = etree.fromstring(xml, parser=etree.XMLParser(resolve_entities=False, no_network=True))
    corpo = raiz.find(".//office:text", NS)
    if corpo is None:
        raise ErroLeitura(f"{caminho.name}: corpo do documento não encontrado.")

    paragrafos = _paragrafos_fora_de_tabela(corpo)
    campos = {}
    faltando = []
    for campo, rotulo in ROTULOS.items():
        v = _busca_rotulo(paragrafos, rotulo)
        if v is None or (v == "" and campo not in OPCIONAIS):
            if campo in OPCIONAIS:
                continue
            faltando.append(rotulo)
        campos[campo] = v
    if faltando:
        raise ErroLeitura(f"{caminho.name}: rótulo(s) não encontrado(s): {', '.join(faltando)}")

    numero = campos["numero_proforma"].replace(" ", "")
    m = re.fullmatch(r"(\d+)(R\d+)?", numero, re.I)
    if not m:
        raise ErroLeitura(f"{caminho.name}: número de proforma inesperado: {numero!r}")

    tabelas = list(corpo.iter(TB + "table"))
    return {
        "arquivo": caminho.name,
        "numero": numero,                 # ex.: 13954R1
        "numero_base": m.group(1),        # ex.: 13954
        "revisao": (m.group(2) or "").upper(),
        "fecha_emision": campos["fecha_emision"],
        "validez": campos.get("validez"),
        "cliente": {
            "nombre": campos["cliente_nombre"],
            "direccion": campos["cliente_direccion"],
            # "DURAN -ECUADOR" -> "DURAN, ECUADOR"
            "ciudad": re.sub(r"\s*-\s*", ", ", campos["cliente_ciudad"]).strip(", "),
            "ruc": campos["cliente_ruc"],
        },
        "condicion_pago": campos["condicion_pago"],
        "fecha_entrega": campos.get("fecha_entrega"),
        "banco": {
            "intermediario": campos.get("banco_intermediario"),
            "beneficiario_banco": campos["banco_beneficiario"],
            "iban": campos["iban"],
            "beneficiario": campos["beneficiario"],
        },
        "itens": _extrai_itens(tabelas),
        "totais": _extrai_totais(tabelas),
    }


def para_json(dados) -> str:
    def conv(o):
        if isinstance(o, Decimal):
            return float(o)
        raise TypeError(type(o))
    return json.dumps(dados, ensure_ascii=False, indent=2, default=conv)


if __name__ == "__main__":
    if len(sys.argv) < 2:
        sys.exit("uso: python leitor_odt.py PROFORMA.odt [...]")
    try:
        saida = [ler_proforma(a) for a in sys.argv[1:]]
    except ErroLeitura as e:
        sys.exit(f"ERRO: {e}")
    print(para_json(saida[0] if len(saida) == 1 else saida))
