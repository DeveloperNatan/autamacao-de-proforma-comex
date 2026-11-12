"""Testes do fluxo completo. Rodar a partir da pasta do projeto:
    .venv/bin/python -m unittest discover -s testes -v
"""
import shutil
import sys
import tempfile
import unittest
import zipfile
from decimal import Decimal
from pathlib import Path

import openpyxl

RAIZ = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(RAIZ))

import gerar  # noqa: E402
from leitor_odt import ErroLeitura, ler_proforma  # noqa: E402

ODT = RAIZ / "testes" / "dados" / "Proforma 13954R1_DISQUISA_Ecuador.odt"


def odt_modificado(destino: Path, trocas: list[tuple[str, str]]) -> Path:
    """Copia o ODT real trocando trechos do content.xml (para casos sintéticos)."""
    with zipfile.ZipFile(ODT) as zin, zipfile.ZipFile(destino, "w") as zout:
        for info in zin.infolist():
            dados = zin.read(info.filename)
            if info.filename == "content.xml":
                xml = dados.decode("utf-8")
                for a, b in trocas:
                    assert a in xml, f"trecho não encontrado no ODT: {a}"
                    xml = xml.replace(a, b)
                dados = xml.encode("utf-8")
            zout.writestr(info, dados)
    return destino


class TestLeitor(unittest.TestCase):
    def test_proforma_13954R1(self):
        pf = ler_proforma(ODT)
        self.assertEqual(pf["numero"], "13954R1")
        self.assertEqual(pf["numero_base"], "13954")
        self.assertEqual(len(pf["itens"]), 9)
        self.assertEqual(pf["totais"]["total_exw"], Decimal("10446.10"))
        self.assertEqual(pf["cliente"]["ruc"], "0990990018001")
        self.assertEqual(pf["condicion_pago"], "100% anticipado a la firma de contrato")
        primeiro = pf["itens"][0]
        self.assertEqual(primeiro["codigo"], "S014030")
        self.assertEqual(primeiro["cantidad"], Decimal("2.00"))
        self.assertEqual(primeiro["valor_unitario"], Decimal("788.20"))
        self.assertFalse(any(i["descripcion"].endswith("-") for i in pf["itens"]))
        self.assertEqual(pf["itens"][-1]["descripcion"], "POLIA MOTOR HU10FHC Ø160MM - 60HZ")

    def test_rotulo_ausente_para(self):
        with tempfile.TemporaryDirectory() as d:
            ruim = odt_modificado(Path(d) / "x.odt", [("RUC/RUT/NIT:", "RUCX:")])
            with self.assertRaisesRegex(ErroLeitura, "RUC/RUT/NIT"):
                ler_proforma(ruim)


class TestGeracao(unittest.TestCase):
    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp())

    def tearDown(self):
        shutil.rmtree(self.tmp)

    def _gera(self, odts, pedido=None):
        args = [str(o) for o in odts] + ["-o", str(self.tmp / "saida.xlsx")]
        if pedido is not None:
            p = self.tmp / "pedido.json"
            p.write_text(__import__("json").dumps(pedido), encoding="utf-8")
            args += ["-p", str(p)]
        return gerar.main(args), self.tmp / "saida.xlsx"

    def test_uma_proforma(self):
        rc, saida = self._gera([ODT], {"fecha_emision": "2026-11-10", "cajas": 3,
                                       "peso_neto_kg": 812.5, "peso_bruto_kg": 950})
        self.assertEqual(rc, 0)
        wb = openpyxl.load_workbook(saida)
        ci, pl = wb["CI"], wb["PL"]
        self.assertNotIn("Planilha1", wb.sheetnames)
        self.assertEqual(ci["G5"].value, "Commercial Invoice: SAL13954/2026")
        self.assertEqual(ci["A25"].value, "Proforma: 13954")
        self.assertEqual(ci["I26"].value, "=H26*C26")
        self.assertEqual(ci["I35"].value, "=SUM(I26:I34)")
        self.assertIn("100% anticipado", str(ci["G11"].value))
        self.assertIn("100% anticipado", str(pl["H20"].value))       # pagamento na PL
        self.assertEqual(ci["B37"].value, "Consignatario / Notificar:")
        self.assertEqual(pl["B38"].value, "Consignatario / Notificar:")  # PL subiu 1 linha
        self.assertEqual(pl["K35"].value, 84835010)
        self.assertEqual(pl["B27"].value, "=CI!B26")
        self.assertIn("812,500", str(ci["H37"].value))
        self.assertEqual(pl["H13"].value, "C. Nº 01/03")

    def test_duas_proformas(self):
        segunda = odt_modificado(self.tmp / "p2.odt", [(">13954<", ">13999<")])
        rc, saida = self._gera([ODT, segunda], {"fecha_emision": "2026-11-10"})
        self.assertEqual(rc, 0)
        wb = openpyxl.load_workbook(saida)
        ci = wb["CI"]
        self.assertEqual(ci["G5"].value, "Commercial Invoice: SAL13954+13999/2026")
        self.assertEqual(ci["A35"].value, "Proforma: 13999")
        self.assertEqual(ci["I45"].value, "=SUM(I26:I44)")
        # rodapé deslocado 10 linhas
        self.assertEqual(ci["B47"].value, "Consignatario / Notificar:")
        self.assertEqual(ci.print_area, "'CI'!$A$1:$I$62")
        assinatura = [i for i in ci._images if i.anchor._from.col == 11 and i.anchor._from.row > 30]
        self.assertEqual(assinatura[0].anchor._from.row, 47)

    def test_condicao_pagamento_divergente(self):
        segunda = odt_modificado(self.tmp / "p2.odt", [
            (">13954<", ">13999<"),
            ("100% anticipado a la firma de contrato", "50% anticipado + 50% contra B/L"),
        ])
        rc, _ = self._gera([ODT, segunda])
        self.assertEqual(rc, 1)
        rc, saida = self._gera([ODT, segunda], {"condicion_pago": "50% + 50%"})
        self.assertEqual(rc, 0)
        self.assertIn("50% + 50%", str(openpyxl.load_workbook(saida)["CI"]["G11"].value))


if __name__ == "__main__":
    unittest.main()
