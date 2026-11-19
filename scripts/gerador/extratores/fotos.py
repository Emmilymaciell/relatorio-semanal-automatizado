# -*- coding: utf-8 -*-
r"""
scripts/gerador/extratores/fotos.py - inventario de fotos por dia/servico e
derivacao das subsecoes 5.2.x a partir do RELATORIO FOTOGRAFICO.

Regra T-17 (CONFIRMADO POR CRUZAMENTO DE FONTES, services_taxonomy.yaml): uma
aba do RELATORIO FOTOGRAFICO <n>.xlsx vira subsecao 5.2.x se, e somente se,
sheet_state == 'visible' E tiver ao menos 1 celula com imagem-na-celula
(marcador `vm="N"` no XML da planilha -- recurso 'Inserir > Imagem na celula'
do Excel 365; openpyxl nao expoe isso via API, por isso lemos o XML cru).

VERIFICADO DIRETAMENTE nesta sessao para RD103/RO_049: abas visiveis+com foto
= TERRAPLENAGEM(9), PAVIMENTAÇÃO(12), DRENAGEM(7), OAE(9); OAC e as demais
ficam ocultas ou sem imagem. Bate 1:1 com as 4 pastas de foto do periodo
(Drenagem, OAE, Pavimentação, Terraplenagem).

NAO faz OCR do overlay da foto (data/UTM/estaca/servico impressos na imagem)
nesta iteracao -- ver report_schema.yaml (`overlay_read: false`,
`status: READY_WITH_HITL`).
"""
import os
import re
import sys
import zipfile
import datetime as dt
import xml.etree.ElementTree as ET

AQUI = os.path.dirname(os.path.abspath(__file__))
SCRIPTS = os.path.dirname(os.path.dirname(AQUI))
if SCRIPTS not in sys.path:
    sys.path.insert(0, SCRIPTS)
import prog_util as U  # noqa: E402
from .. import assert_nao_e_golden  # noqa: E402

_NS_MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
_NS_REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"


def inventariar_pasta(pasta_fotos_raiz, dias_iso):
    """Para cada dia do periodo, lista as subpastas de servico e a contagem de
    arquivos. Aceita pastas de dia no formato 'AA-MM-DD' OU 'AA_MM_DD'."""
    resultado = []
    for diso in dias_iso:
        d = dt.date.fromisoformat(diso)
        alvo_h = "%02d-%02d-%02d" % (d.year % 100, d.month, d.day)
        alvo_u = alvo_h.replace("-", "_")
        pasta_dia = None
        if pasta_fotos_raiz and U.exists(pasta_fotos_raiz):
            for nome in U.listdir(pasta_fotos_raiz):
                if nome in (alvo_h, alvo_u):
                    pasta_dia = os.path.join(pasta_fotos_raiz, nome)
                    break
        if not pasta_dia:
            resultado.append({"date": diso, "status": "NOT_FOUND", "servicos": [], "total": 0})
            continue
        servicos = []
        for servico in U.listdir(pasta_dia):
            sp = os.path.join(pasta_dia, servico)
            if not U.isdir(sp):
                continue
            arquivos = [f for f in U.listdir(sp) if not U.isdir(os.path.join(sp, f))]
            servicos.append({"pasta": servico, "n": len(arquivos), "arquivos": arquivos})
        resultado.append({"date": diso, "status": "OK", "servicos": servicos,
                          "total": sum(s["n"] for s in servicos)})
    return resultado


def abas_com_foto(caminho_relatorio_fotografico):
    """Devolve {status, abas:[{nome,estado,n_fotos}], subsecoes_5_2:[nomes]}."""
    if not caminho_relatorio_fotografico:
        return {"status": "NOT_FOUND"}
    assert_nao_e_golden(caminho_relatorio_fotografico)
    if not U.exists(caminho_relatorio_fotografico):
        return {"status": "NOT_FOUND", "path": caminho_relatorio_fotografico}

    try:
        local = U.copiar_para_tmp(
            caminho_relatorio_fotografico, "relatorio_fotografico",
            "relfoto" + os.path.splitext(caminho_relatorio_fotografico)[1])
    except OSError as e:
        return {"status": "UNREADABLE", "path": caminho_relatorio_fotografico, "erro": str(e)}

    try:
        z = zipfile.ZipFile(local)
        wbxml = ET.fromstring(z.read("xl/workbook.xml"))
        sheets = [(s.get("name"), s.get(_NS_REL + "id"), s.get("state", "visible"))
                  for s in wbxml.find(_NS_MAIN + "sheets")]
        rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        relmap = {r.get("Id"): r.get("Target") for r in rels}
        abas = []
        for nome, rid, estado in sheets:
            target = relmap.get(rid, "")
            parte = "xl/" + target if not target.startswith("/") else target.lstrip("/")
            try:
                dados = z.read(parte)
            except KeyError:
                abas.append({"nome": nome, "estado": estado, "n_fotos": 0})
                continue
            n_vm = len(re.findall(rb' vm="\d+"', dados))
            abas.append({"nome": nome, "estado": estado, "n_fotos": n_vm})
    except (KeyError, zipfile.BadZipFile, ET.ParseError, OSError) as e:
        return {"status": "UNREADABLE", "path": caminho_relatorio_fotografico, "erro": str(e)}

    subsecoes = [a["nome"] for a in abas if a["estado"] == "visible" and a["n_fotos"] > 0]
    return {"status": "OK", "path": caminho_relatorio_fotografico, "abas": abas,
            "subsecoes_5_2": subsecoes}
