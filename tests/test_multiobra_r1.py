# -*- coding: utf-8 -*-
"""
test_multiobra_r1.py - um caso por bug GLOBAL_ENGINE_BUG da rodada multiobra-r1
(docs/multiobra/PREPARACAO_MULTIOBRA.md). Sem SharePoint: fixtures minimos em
pasta temporaria / memoria.

Execute:  python tests/test_multiobra_r1.py
"""
import os
import sys
import tempfile
import datetime as dt

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(os.path.dirname(AQUI), "scripts"))
import prog_util as U  # noqa: E402
from gerador import config as C  # noqa: E402
from gerador import source_discovery as SD  # noqa: E402
from gerador.extratores import nc as EX_NC  # noqa: E402
from gerador.extratores import parecer as EX_PAR  # noqa: E402
from gerador.figuras import ficha_acumulada as FA  # noqa: E402
from gerador import content_builder as CB  # noqa: E402

import openpyxl  # noqa: E402
from openpyxl.worksheet.table import Table  # noqa: E402

falhas = []


def check(nome, cond, detalhe=""):
    print(("OK   " if cond else "FAIL ") + nome + ((" -- " + detalhe) if detalhe else ""))
    if not cond:
        falhas.append(nome)


def _touch(*partes):
    p = os.path.join(*partes)
    os.makedirs(os.path.dirname(p), exist_ok=True)
    open(p, "wb").write(b"x")
    return p


def main():
    tmp = tempfile.mkdtemp(prefix="prog_r1_")

    # ---- Bug 1: NC-022-1 (tracker) x NC-022 (nome do PDF), e vice-versa
    idx = {}
    orig_base = U.BASE_OBRAS
    try:
        U.BASE_OBRAS = tmp
        pasta = os.path.join(tmp, "O", "3_Gestao Tecnica de Obra", "3. Notificações", "Not. Geradas")
        a = _touch(pasta, "NC-022 - Notificacao.pdf")
        b = _touch(pasta, "NC-047-1 - Notificacao.pdf")
        idx = EX_NC.indexar_pdfs("O")
    finally:
        U.BASE_OBRAS = orig_base
    check("bug1 NC-022-1 acha PDF 'NC-022 ...' (prefixo base)",
          EX_NC.pdfs_da_nc(idx, "NC-022-1") == [a], str(idx))
    check("bug1 NC-047-1 acha PDF 'NC-047-1 ...' (chave completa)",
          EX_NC.pdfs_da_nc(idx, "NC-047-1") == [b])
    check("bug1 documento inexistente -> []", EX_NC.pdfs_da_nc(idx, "NC-999-1") == [])

    # ---- Bug 2: rodovia 'RD-102 A' (registry) x 'RD-102A' (tracker)
    check("bug2 norm_rodovia colide", C.norm_rodovia("RD-102 A") == C.norm_rodovia("RD-102A")
          == C.norm_rodovia("rd 102a"))
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "VISAO_GERAL"
    cab = ["Documento", "Rodovia", "Data de Abertura", "Status ", "Irregularidade ", "Resumo", "Natureza"]
    ws.append(cab)
    ws.append(["NC-047-1", "RD-102A", dt.datetime(2026, 9, 1), "Aberta", "Falha de teste.", "r", "Execucao"])
    t = Table(displayName="Tbl_VISAO", ref="A1:G2")
    ws.add_table(t)
    xl = os.path.join(tmp, "tracker.xlsx")
    wb.save(xl)
    orig_cp = U.copiar_para_tmp
    U.copiar_para_tmp = lambda origem, *a, **k: origem
    try:
        r = EX_NC.ler(xl, "RD-102 A", "2026-08-29", "2026-09-04")
    finally:
        U.copiar_para_tmp = orig_cp
    check("bug2 ler() acha a NC-047-1 da RD-102 A",
          r.get("status") == "OK" and [i["documento"] for i in r["issued_in_period"]] == ["NC-047-1"],
          str(r.get("issued_in_period")))

    # ---- Bug 3: pasta de ata '8. Atas de Reunioes' (RD101/RD102A) e '9. Ata de Reuniao' (RD104)
    rx = SD._rx(C.source_layout(), "pasta_ata_reuniao")
    for nome in ("8. Atas de Reuniões", "9. Ata de Reunião", "9. Atas de Reuniao"):
        base = os.path.join(tmp, "gto_%d" % len(nome), nome)
        d = os.path.dirname(base)
        os.makedirs(base, exist_ok=True)
        os.makedirs(os.path.join(d, "6. Acompanhamento Ambiental"), exist_ok=True)
        ach = C.resolver_subpasta(d, prefixo="9. ata de reuni", regex=rx)
        check("bug3 ata '%s'" % nome, ach is not None and os.path.basename(ach) == nome)

    # ---- Bug 4: '1º de setembro de 2026'
    s = EX_PAR.extrair_sinais("Foram realizadas verificações em campo no dia 1º de setembro de 2026 na RD-102 A.")
    check("bug4 parecer reconhece '1º de setembro'",
          [x["data"] for x in s["datas_campo"]] == ["2026-09-01"], str(s["datas_campo"]))
    s = EX_PAR.extrair_sinais("verificações em campo em 29 de agosto de 2026")
    check("bug4 sem regressao '29 de agosto'", [x["data"] for x in s["datas_campo"]] == ["2026-08-29"])

    # ---- Bug 5: typo 'Dário de obra digital assinado'
    diario = os.path.join(tmp, "4. Diario de obras")
    for nome in ("Dário de obra digital assinado", "Diário de obra digital assinado", "Diário de obra físico"):
        os.makedirs(os.path.join(diario, nome), exist_ok=True)
    rx5 = SD._rx(C.source_layout(), "pasta_diario_obras", "subpastas", "assinado")
    d5 = os.path.join(tmp, "d5")
    os.makedirs(os.path.join(d5, "Dário de obra digital assinado"))
    ach = C.resolver_subpasta(d5, prefixo="diario de obra digital assinado", regex=rx5)
    check("bug5 typo 'Dário' aceito", ach is not None)
    d5b = os.path.join(tmp, "d5b")
    os.makedirs(os.path.join(d5b, "Diário de obra físico"))
    check("bug5 'físico' nao casa com assinado",
          C.resolver_subpasta(d5b, prefixo="diario de obra digital assinado", regex=rx5) is None)

    # ---- Bug 6: 'Cópia de Ficha de Campo Acumulada'
    aux = os.path.join(tmp, "aux")
    _touch(aux, "Cópia de Ficha de Campo Acumulada RD-102A.xlsx")
    rx6 = SD._rx(C.source_layout(), "ficha_acumulada", raiz="auxiliares")
    ach = C.resolver_arquivo(aux, prefixo="ficha de campo acumulada", extensoes=[".xlsx"], regex=rx6)
    check("bug6 'Cópia de ...' aceita", ach is not None)
    aux2 = os.path.join(tmp, "aux2")
    _touch(aux2, "Outro arquivo.xlsx")
    check("bug6 outro arquivo nao casa",
          C.resolver_arquivo(aux2, prefixo="ficha de campo acumulada", extensoes=[".xlsx"], regex=rx6) is None)

    # ---- Bug 7: cabecalhos 'Estaca i' / 'Estaca f'
    wb7 = openpyxl.Workbook()
    ws7 = wb7.active
    for c, v in enumerate(["Estaca i", None, None, "Estaca f", None, None, "Lado", "Serviço"], start=1):
        ws7.cell(3, c, v)
    r7, cm = FA._achar_cabecalho(ws7)
    check("bug7 'Estaca i'/'Estaca f' reconhecidos",
          r7 == 3 and cm.get("estaca inicial") == 1 and cm.get("estaca final") == 4, str(cm))
    wb8 = openpyxl.Workbook()
    ws8 = wb8.active
    for c, v in enumerate(["Estaca Inicial", None, None, "Estaca Final", None, None, "Serviço"], start=1):
        ws8.cell(2, c, v)
    r8, cm8 = FA._achar_cabecalho(ws8)
    check("bug7 sem regressao 'Estaca Inicial'", r8 == 2 and cm8.get("estaca final") == 4)

    # ---- Bug 8: concordancia singular/plural
    um = CB._fmt_n("s8_1_positiva_intro", 1)
    dois = CB._fmt_n("s8_1_positiva_intro", 2)
    check("bug8 singular n=1", "foi emitida 1 não conformidade pelo Programa, anexada" in um, um)
    check("bug8 plural n=2 inalterado", "foram emitidas 2 não conformidades pelo Programa, anexadas" in dois, dois)

    print("=" * 60)
    if falhas:
        print("FALHAS:", falhas)
        return 1
    print("RESULTADO: OK")
    return 0


if __name__ == "__main__":
    sys.exit(main())
