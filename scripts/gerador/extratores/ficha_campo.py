# -*- coding: utf-8 -*-
r"""
scripts/gerador/extratores/ficha_campo.py - clima e tabela de servicos/estacas
da Ficha de Campo diaria (aba 'FICHA CAMPO').

VERIFICADO DIRETAMENTE nesta sessao (2026-09-17) contra o arquivo real
'2026-09-04 Ficha de Campo RD 103.xlsx':

  - clima: rotulos em N16/O16/P16 (Bom/Chuvoso/Impraticavel), marca 'X' na
    linha 17. Bate com a spec (FASE3_SLICE_1_2_SPEC.md Sec.3) e com
    docs/mapeamento/06_fontes_dados.md.

  - tabela de servicos: cabecalho na LINHA 21 (nao 24/25 como diz a spec),
    dados a partir da LINHA 22. Colunas (verificadas por rotulo, nao por
    letra fixa): B/C/D=estaca inicial, E/F/G=estaca final, H=lado,
    J=extensao(km), K=servico, M=situacao, O=foto, Q=descricao, X=observacao.
    Isso bate com docs/mapeamento/06_fontes_dados.md (mais antigo) e DIVERGE
    da spec nova (que diz I=extensao, J=servico, L=situacao, M=foto,
    O=descricao, AC=observacao). Ver aviso em config/source_layout.yaml.

  Este modulo NAO fixa letras de coluna: localiza o cabecalho varrendo a
  planilha por texto normalizado ('estaca i', 'servico', ...), o que o torna
  robusto a pequenos deslocamentos entre obras/semanas.
"""
import os
import sys

import openpyxl
from openpyxl.utils import get_column_letter

AQUI = os.path.dirname(os.path.abspath(__file__))
SCRIPTS = os.path.dirname(os.path.dirname(AQUI))
if SCRIPTS not in sys.path:
    sys.path.insert(0, SCRIPTS)
import prog_util as U  # noqa: E402
from .. import assert_nao_e_golden  # noqa: E402

ABA = "FICHA CAMPO"
_ROTULOS_CLIMA = {"bom": "Bom", "chuvoso": "Chuvoso", "impraticavel": "Impraticável"}

_ROTULOS_TABELA = {
    "estaca i": "estaca_inicial", "estaca f": "estaca_final", "lado": "lado",
    "extensao (km)": "extensao_km", "servico": "servico_codigo",
    "situacao": "situacao", "foto": "foto", "descricao": "descricao",
    "observacao": "observacao",
}


def _n(v):
    return U.norm(v) if v is not None else ""


def _achar_linha_clima(ws, varredura=25):
    for r in range(1, varredura + 1):
        achados = {}
        for c in range(1, 40):
            v = _n(ws.cell(r, c).value)
            if v in _ROTULOS_CLIMA:
                achados[c] = v
        if len(achados) >= 2:
            return r, achados
    return None, {}


def extrair_clima(ws):
    linha_rotulos, cols = _achar_linha_clima(ws)
    if linha_rotulos is None:
        return None, {"status": "NOT_FOUND",
                       "motivo": "linha com rotulos Bom/Chuvoso/Impraticavel nao encontrada"}
    linha_marca = linha_rotulos + 1
    marcados = []
    for c, rot in cols.items():
        v = ws.cell(linha_marca, c).value
        if v is not None and str(v).strip().upper() == "X":
            marcados.append((rot, "%s%d" % (get_column_letter(c), linha_marca)))
    if len(marcados) == 0:
        return None, {"status": "NOT_FOUND",
                       "motivo": "nenhum 'X' na linha %d" % linha_marca,
                       "linha_rotulos": linha_rotulos}
    if len(marcados) > 1:
        return None, {"status": "AMBIGUOUS",
                       "motivo": "mais de um 'X' na linha %d: %s" % (linha_marca, marcados),
                       "linha_rotulos": linha_rotulos}
    rot, cel = marcados[0]
    return _ROTULOS_CLIMA[rot], {"status": "OK", "celula": cel, "linha_rotulos": linha_rotulos}


def _achar_cabecalho_tabela(ws, varredura=40):
    for r in range(1, varredura + 1):
        for c in range(1, 60):
            if _n(ws.cell(r, c).value) == "estaca i":
                return r
    return None


def extrair_tabela_servicos(ws):
    header_row = _achar_cabecalho_tabela(ws)
    if header_row is None:
        return [], {"status": "NOT_FOUND", "motivo": "cabecalho 'Estaca i' nao encontrado"}

    mapa_col = {}
    for c in range(1, 60):
        v = _n(ws.cell(header_row, c).value)
        if v in _ROTULOS_TABELA and v not in mapa_col:
            mapa_col[v] = c

    c_ei, c_ef = mapa_col.get("estaca i"), mapa_col.get("estaca f")
    c_serv = mapa_col.get("servico")
    if not c_ei or not c_serv:
        return [], {"status": "NOT_FOUND",
                     "motivo": "colunas essenciais (estaca i / servico) ausentes no cabecalho",
                     "linha_cabecalho": header_row}

    def estaca(c_base):
        if not c_base:
            return None
        partes = [ws.cell(r_atual, c_base + i).value for i in range(3)]
        rotulo = "".join(str(p) for p in partes if p not in (None, ""))
        return {"km": partes[0], "separador": partes[1], "m": partes[2], "rotulo": rotulo}

    linhas = []
    r_atual = header_row + 1
    while True:
        serv = ws.cell(r_atual, c_serv).value
        if serv in (None, ""):
            break
        linhas.append({
            "linha": r_atual,
            "estaca_inicial": estaca(c_ei),
            "estaca_final": estaca(c_ef),
            "lado": ws.cell(r_atual, mapa_col["lado"]).value if "lado" in mapa_col else None,
            "extensao_km": ws.cell(r_atual, mapa_col["extensao (km)"]).value
                           if "extensao (km)" in mapa_col else None,
            "servico_codigo": str(serv).strip(),
            "situacao": ws.cell(r_atual, mapa_col["situacao"]).value
                        if "situacao" in mapa_col else None,
            "foto": ws.cell(r_atual, mapa_col["foto"]).value if "foto" in mapa_col else None,
            "descricao": ws.cell(r_atual, mapa_col["descricao"]).value
                         if "descricao" in mapa_col else None,
            "observacao": ws.cell(r_atual, mapa_col["observacao"]).value
                          if "observacao" in mapa_col else None,
        })
        r_atual += 1

    return linhas, {"status": "OK", "linha_cabecalho": header_row,
                     "linha_dados_inicial": header_row + 1, "n_linhas": len(linhas),
                     "colunas_usadas": {k: get_column_letter(v) for k, v in mapa_col.items()}}


def ler(caminho):
    """Le clima + tabela de servicos de UMA Ficha de Campo diaria."""
    if not caminho:
        return {"status": "NOT_FOUND"}
    assert_nao_e_golden(caminho)
    if not U.exists(caminho):
        return {"status": "NOT_FOUND", "path": caminho}

    try:
        destino = "ficha_%s%s" % (U.slug(os.path.basename(caminho)),
                                   os.path.splitext(caminho)[1])
        local = U.copiar_para_tmp(caminho, "ficha_campo", destino)
        wb = openpyxl.load_workbook(local, data_only=True)
    except OSError as e:
        return {"status": "UNREADABLE", "path": caminho, "erro": str(e)}

    if ABA not in wb.sheetnames:
        return {"status": "UNREADABLE", "path": caminho,
                "erro": "aba '%s' ausente (abas: %s)" % (ABA, wb.sheetnames)}

    ws = wb[ABA]
    clima, ev_clima = extrair_clima(ws)
    servicos, ev_serv = extrair_tabela_servicos(ws)

    return {
        "status": "OK", "path": caminho,
        "clima": clima, "clima_evidencia": ev_clima,
        "servicos": servicos, "servicos_evidencia": ev_serv,
    }
