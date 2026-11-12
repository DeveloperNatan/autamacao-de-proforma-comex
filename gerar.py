"""Gera a Commercial Invoice + Packing List (XLSX Satake) a partir de proformas .odt.

Uso:
  python gerar.py PROFORMA.odt [PROFORMA2.odt ...] [-p pedido.json] [-o saida.xlsx] [--json dados.json]

Ordem: ler ODTs -> validar cada proforma -> consolidar -> montar contexto ->
escrever XLSX -> validar o XLSX gerado.
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from datetime import date
from decimal import Decimal
from pathlib import Path

from escritor_xlsx import ErroEscrita, escrever
from leitor_odt import ErroLeitura, ler_proforma, para_json
from validador import ErroValidacao, validar_proforma, validar_xlsx

BASE = Path(__file__).resolve().parent
MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio",
         "agosto", "septiembre", "octubre", "noviembre", "diciembre"]


class ErroConsolidacao(Exception):
    pass


def _json(caminho: Path, padrao=None):
    if not caminho.exists():
        if padrao is not None:
            return padrao
        raise FileNotFoundError(caminho)
    with open(caminho, encoding="utf-8") as f:
        dados = json.load(f)
    return {k: v for k, v in dados.items() if not k.startswith("_")} if isinstance(dados, dict) else dados


def fmt_br(valor, casas: int) -> str:
    s = f"{Decimal(str(valor)):,.{casas}f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def _unico(proformas, chave, rotulo, override=None):
    if override not in (None, ""):
        return override
    valores = {json.dumps(chave(p), sort_keys=True, ensure_ascii=False) for p in proformas}
    if len(valores) > 1:
        detalhes = "\n".join(f"  {p['numero']}: {chave(p)}" for p in proformas)
        raise ErroConsolidacao(
            f"As proformas têm '{rotulo}' diferentes:\n{detalhes}\n"
            f"Defina o valor no pedido.json para seguir."
        )
    return chave(proformas[0])


def montar_contexto(proformas: list[dict], pedido: dict, padroes: dict, clientes: dict,
                    traducoes: dict) -> tuple[dict, list[dict], list[str]]:
    avisos: list[str] = []
    nums = [p["numero_base"] for p in proformas]
    if len(set(nums)) != len(nums):
        raise ErroConsolidacao(f"Proforma repetida na lista: {nums}")

    ruc = _unico(proformas, lambda p: p["cliente"]["ruc"], "RUC do cliente")
    banco = _unico(proformas, lambda p: p["banco"], "Información bancaria")
    cond = _unico(proformas, lambda p: p["condicion_pago"], "Condición de pago",
                  pedido.get("condicion_pago"))
    if pedido.get("condicion_pago") and any(p["condicion_pago"] != cond for p in proformas):
        avisos.append(f"Condición de pago sobrescrita pelo pedido.json: '{cond}' "
                      f"(proforma: '{proformas[0]['condicion_pago']}')")

    cad = clientes.get(ruc)
    if cad is None:
        avisos.append(f"Cliente RUC {ruc} não está em config/clientes.json (Attn./Fono ficam pendentes)")
        cad = {}
    cliente = {**proformas[0]["cliente"]}
    for k in ("nombre", "direccion", "ciudad"):  # cadastro pode trazer a grafia preferida
        if cad.get(k):
            cliente[k] = cad[k]
    for k in ("attn", "fono", "email"):
        cliente[k] = cad.get(k)

    consig = pedido.get("consignatario") or cad.get("consignatario") or cliente

    manter_rev = pedido.get("manter_revisao", padroes.get("manter_revisao", False))
    exib = [p["numero"] if manter_rev else p["numero_base"] for p in proformas]

    d = date.fromisoformat(pedido["fecha_emision"]) if pedido.get("fecha_emision") else date.today()
    compuesto = f"{padroes.get('prefixo_invoice', 'SAL')}{'+'.join(exib)}/{d.year}"

    destino = pedido.get("destino") or cad.get("destino")
    destino_marca = pedido.get("destino_marca") or (destino and re.sub(r"\s*-\s*", ", ", destino))

    # ---- itens / packing ----
    itens_pl = pedido.get("itens_pl", {})
    blocos, sem_trad = [], []
    tot_neto = tot_bruto = Decimal(0)
    tem_peso_item = False
    for p, num_exib in zip(proformas, exib):
        pl_pf = itens_pl.get(p["numero_base"], {}) or itens_pl.get(p["numero"], {})
        codigos = [i["codigo"] for i in p["itens"]]
        for cod in pl_pf:
            if cod not in codigos:
                raise ErroConsolidacao(f"itens_pl de {p['numero_base']} cita o código {cod}, que não está na proforma")
        itens = []
        for it in p["itens"]:
            desc = it["descripcion"]
            if pedido.get("traducir", True):
                if it["codigo"] in traducoes:
                    desc = traducoes[it["codigo"]]
                else:
                    sem_trad.append(it["codigo"])
            pl = pl_pf.get(it["codigo"], {})
            if pl.get("peso_neto") is not None:
                tem_peso_item = True
                tot_neto += Decimal(str(pl["peso_neto"]))
            if pl.get("peso_bruto") is not None:
                tot_bruto += Decimal(str(pl["peso_bruto"]))
            itens.append({**it, "descripcion": desc, "ncm": int(it["ncm"]), "pl": pl})
        blocos.append({"numero": p["numero"], "numero_exibido": num_exib, "itens": itens})

    if sem_trad and pedido.get("traducir", True):
        avisos.append(f"{len(sem_trad)} descrição(ões) sem tradução em config/traducciones_es.json "
                      f"(mantidas como na proforma): {', '.join(sem_trad)}")

    peso_neto = pedido.get("peso_neto_kg")
    peso_bruto = pedido.get("peso_bruto_kg")
    if peso_neto is None and tem_peso_item:
        peso_neto = tot_neto
    if peso_bruto is None and tem_peso_item:
        peso_bruto = tot_bruto
    cajas = pedido.get("cajas")

    ctx = {
        "numero_invoice": pedido.get("numero_invoice") or compuesto,
        "proformas_compuesto": compuesto,
        "lugar_fecha": f"{padroes['lugar_emision']}, {d.day} de {MESES[d.month - 1]}, {d.year}",
        "origen": padroes["origen"],
        "moneda": padroes["moneda"],
        "marca": padroes["marca"],
        "empresa": padroes["empresa"],
        "destino": destino,
        "destino_marca": destino_marca,
        "via_embarque": pedido.get("via_embarque"),
        "condicion_pago": cond,
        "cliente": cliente,
        "consignatario": consig,
        "banco": banco,
        "peso_neto": fmt_br(peso_neto, 3) if peso_neto is not None else None,
        "peso_bruto": fmt_br(peso_bruto, 2) if peso_bruto is not None else None,
        "cajas": cajas,
        "cajas_2d": f"{int(cajas):02d}" if cajas else None,
    }
    return ctx, blocos, avisos


def nome_saida(proformas, ctx) -> str:
    nums = "+".join(p["numero_base"] for p in proformas)
    nome = re.sub(r"[^\w]", "", ctx["cliente"]["nombre"].split()[0]).upper()
    pais = re.sub(r"[^\w]", "", ctx["cliente"]["ciudad"].split(",")[-1]).upper()
    return f"CI+PL{nums} {nome}_{pais}_SAL.xlsx"


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="Gera CI + PL (XLSX Satake) a partir de proformas .odt")
    ap.add_argument("odts", nargs="+", help="uma ou mais proformas .odt (na ordem da invoice)")
    ap.add_argument("-p", "--pedido", help="JSON com os campos manuais desta invoice")
    ap.add_argument("-o", "--saida", help="arquivo XLSX de saída (padrão: saida/CI+PL<nums> ...xlsx)")
    ap.add_argument("--json", help="também grava o JSON extraído das proformas neste caminho")
    a = ap.parse_args(argv)

    try:
        proformas = [ler_proforma(o) for o in a.odts]
        for p in proformas:
            total = validar_proforma(p)
            print(f"OK  {p['arquivo']}: {len(p['itens'])} itens, Total EXW USD {fmt_br(total, 2)}")
        if a.json:
            Path(a.json).write_text(para_json(proformas), encoding="utf-8")

        mapeamento = _json(BASE / "config" / "mapeamento.json")
        padroes = _json(BASE / "config" / "padroes.json")
        clientes = _json(BASE / "config" / "clientes.json", {})
        traducoes = _json(BASE / "config" / "traducciones_es.json", {})
        pedido = _json(Path(a.pedido)) if a.pedido else {}

        ctx, blocos, avisos = montar_contexto(proformas, pedido, padroes, clientes, traducoes)
        destino = Path(a.saida) if a.saida else BASE / "saida" / nome_saida(proformas, ctx)
        destino.parent.mkdir(parents=True, exist_ok=True)

        escrever(mapeamento, ctx, blocos, destino, BASE)
        avisos += [f"pendente  {x}" for x in validar_xlsx(destino, mapeamento["layout"], proformas)]
    except (ErroLeitura, ErroValidacao, ErroConsolidacao, ErroEscrita, FileNotFoundError) as e:
        print(f"ERRO: {e}", file=sys.stderr)
        return 1

    total = sum(p["totais"]["total_exw"] for p in proformas)
    print(f"OK  XLSX validado: total CI = Total EXW = USD {fmt_br(total, 2)}")
    print(f"Gerado: {destino}")
    for av in avisos:
        print(f"AVISO {av}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
