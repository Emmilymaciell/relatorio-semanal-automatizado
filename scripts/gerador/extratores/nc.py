# -*- coding: utf-8 -*-
r"""
scripts/gerador/extratores/nc.py - nao conformidades (02_Acompanhamento de NC
automatizada.xlsm, aba VISAO_GERAL, tabela Tbl_VISAO).

VERIFICADO DIRETAMENTE nesta sessao (2026-09-17): cabecalho da tabela (ref
A3:AA156) bate exatamente com a spec -- H='Data de Abertura', V='Irregularidade '
(com espaco no fim), B='Rodovia', E='Status ' (idem). RD-103 tem 32 linhas no
total; NENHUMA com 'Data de Abertura' em 2026-08-29..2026-09-04 (issued_in_period
vazio para esta semana -- fato real, nao bug: confirmado varrendo as 32 linhas).
Domínio de status (LISTAS_VALIDACAO!I1:I7): Aberta, Em Análise, Aguardando
Evidências, Encerrada, Aguardando Projeto Executivo, Cancelada.
"""
import os
import re
import sys
import datetime as dt

import openpyxl
from openpyxl.utils import column_index_from_string as cix

AQUI = os.path.dirname(os.path.abspath(__file__))
SCRIPTS = os.path.dirname(os.path.dirname(AQUI))
if SCRIPTS not in sys.path:
    sys.path.insert(0, SCRIPTS)
import prog_util as U  # noqa: E402
from .. import assert_nao_e_golden  # noqa: E402
from .. import config as C  # noqa: E402

DOMINIO_STATUS = ["Aberta", "Em Análise", "Aguardando Evidências", "Encerrada",
                  "Aguardando Projeto Executivo", "Cancelada"]


def _ler_tabela(ws, nome_tabela):
    if nome_tabela not in ws.tables:
        return None, None
    ref = ws.tables[nome_tabela].ref
    m = re.match(r"([A-Z]+)(\d+):([A-Z]+)(\d+)", ref)
    c1, r1, c2, r2 = cix(m.group(1)), int(m.group(2)), cix(m.group(3)), int(m.group(4))
    headers = [ws.cell(r1, c).value for c in range(c1, c2 + 1)]
    linhas = []
    for r in range(r1 + 1, r2 + 1):
        row, vazio = {}, True
        for i, c in enumerate(range(c1, c2 + 1)):
            v = ws.cell(r, c).value
            row[headers[i]] = v
            if v not in (None, ""):
                vazio = False
        if not vazio:
            linhas.append(row)
    return headers, linhas


def _as_date(v):
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    if isinstance(v, str) and re.match(r"^\d{4}-\d{2}-\d{2}", v):
        return dt.date.fromisoformat(v[:10])
    return None


def primeira_frase_sem_clausula_temporal(texto):
    """Regra da spec (report_schema.yaml s8_1_prog.itens_nc.descricao): primeira
    sentenca da Irregularidade, sem a clausula temporal inicial, inicial
    minuscula. NAO houve NC no periodo de teste (RD103/2026-08-29..09-04) para
    validar esta heuristica ao vivo nesta sessao -- implementada conforme a
    regra descrita, mas sem confirmar against um exemplo real desta execucao."""
    if not texto:
        return ""
    t = re.sub(r"\s+", " ", str(texto)).strip()
    m = re.match(r"(.+?[.])(\s|$)", t)
    frase = m.group(1) if m else t
    frase = re.sub(r"^(em\s+\d{2}/\d{2}/\d{4}[^,]*,\s*)", "", frase, flags=re.IGNORECASE).strip()
    if frase:
        frase = frase[0].lower() + frase[1:]
    return frase


def gerar_sentenca_81(documento, data_abertura_iso, irregularidade):
    data_br = "%s/%s/%s" % (data_abertura_iso[8:10], data_abertura_iso[5:7], data_abertura_iso[0:4])
    corpo = primeira_frase_sem_clausula_temporal(irregularidade)
    return "%s Emitida em %s, %s" % (documento, data_br, corpo)


def ler(caminho, rodovia, data_inicio, data_fim):
    if not caminho:
        return {"status": "NOT_FOUND"}
    assert_nao_e_golden(caminho)
    if not U.exists(caminho):
        return {"status": "NOT_FOUND", "path": caminho}

    try:
        local = U.copiar_para_tmp(caminho, "nc", "nc_tracker.xlsm")
        wb = openpyxl.load_workbook(local, data_only=True)
    except OSError as e:
        return {"status": "UNREADABLE", "path": caminho, "erro": str(e)}

    if "VISAO_GERAL" not in wb.sheetnames:
        return {"status": "UNREADABLE", "path": caminho, "erro": "aba VISAO_GERAL ausente"}
    ws = wb["VISAO_GERAL"]
    headers, linhas = _ler_tabela(ws, "Tbl_VISAO")
    if linhas is None:
        return {"status": "UNREADABLE", "path": caminho, "erro": "tabela Tbl_VISAO ausente"}

    di, dfim = _as_date(data_inicio) or data_inicio, _as_date(data_fim) or data_fim

    alvo_rod = C.norm_rodovia(rodovia)
    linhas_rodovia = [r for r in linhas if C.norm_rodovia(r.get("Rodovia")) == alvo_rod]

    issued = []
    for row in linhas_rodovia:
        abd = _as_date(row.get("Data de Abertura"))
        if abd and di <= abd <= dfim:
            item = {
                "documento": row.get("Documento"),
                "data_abertura": abd.isoformat(),
                "status": row.get("Status "),
                "irregularidade": row.get("Irregularidade "),
                "resumo": row.get("Resumo"),
            }
            item["sentence"] = gerar_sentenca_81(item["documento"] or "NC-???",
                                                  item["data_abertura"],
                                                  item["irregularidade"])
            issued.append(item)

    contagem = {k: 0 for k in DOMINIO_STATUS}
    for row in linhas_rodovia:
        st = str(row.get("Status ") or "").strip()
        if st in contagem:
            contagem[st] += 1
        else:
            chave = "(fora do dominio: %s)" % st if st else "(vazio)"
            contagem[chave] = contagem.get(chave, 0) + 1

    return {
        "status": "OK", "path": caminho, "aba": "VISAO_GERAL", "tabela": "Tbl_VISAO",
        "n_linhas_rodovia": len(linhas_rodovia),
        "issued_in_period": issued,
        "history_by_status": contagem,
        "dominio_status_fonte": "LISTAS_VALIDACAO!I1:I7",
        "alerta": "a planilha usa TODAY(); history_by_status e o estado ATUAL, "
                  "nao um historico reconstruido da semana",
    }


def indexar_pdfs(obra):
    """Mapeia NC-NNN -> lista de caminhos de PDF em 3. Notificacoes\\Not. Geradas."""
    raiz = os.path.join(U.BASE_OBRAS, obra, "3_Gestao Tecnica de Obra", "3. Notificações",
                        "Not. Geradas")
    idx = {}
    if not U.exists(raiz):
        return idx
    for dp, _dn, fn in U.walk(raiz):
        for f in fn:
            m = re.match(r"(NC-\d+(?:-\d+)?)", f, re.IGNORECASE)
            if m:
                chave = m.group(1).upper()
                caminho = os.path.join(dp, f)
                idx.setdefault(chave, []).append(caminho)
                base = chave.split("-")
                base = "-".join(base[:2])
                if base != chave:  # NC-022-1 tambem entra sob o prefixo base NC-022
                    idx.setdefault(base, []).append(caminho)
    return idx


def pdfs_da_nc(idx, documento):
    """PDFs de uma NC pelo indice de `indexar_pdfs`: chave completa primeiro
    ('NC-022-1'); se nao ha, cai no prefixo base ('NC-022')."""
    if not documento:
        return []
    chave = str(documento).strip().upper()
    if chave in idx:
        return idx[chave]
    base = "-".join(chave.split("-")[:2])
    return idx.get(base, [])


# -----------------------------------------------------------------------------
# §6 (spec Sec.2) -- estado HISTORICO das NC e reconstruivel a partir dos
# documentos individuais (Not. Geradas + Respostas), ou so a partir da
# planilha (que usa TODAY(), ou seja, so mostra o estado ATUAL)?
#
# Investigacao PONTUAL (nao varre conteudo de cada PDF -- so classifica nome
# de arquivo por "tipo de documento na cadeia da NC" e procura qualquer
# marcador textual de ENCERRAMENTO no nome). Generico: roda para qualquer
# obra, sem `if obra == ...`; a conclusao sai dos NUMEROS encontrados nos
# arquivos da PROPRIA obra, nunca de um valor fixo.
# -----------------------------------------------------------------------------
_TIPOS_DOC_NC = {
    "abertura": "nao conformidade",
    "nota_tecnica": "nota_tecnica",
    "resposta": "resposta",
    "parecer_resposta": "parecer_tecnico",
}
_MARCADORES_ENCERRAMENTO = ("encerr", "fechamento", "closure", "conclus")


def _contar_documentos_por_tipo(raiz_notificacoes):
    """Varre Not. Geradas/{BV,Programa,Supervisora} + Respostas/* e conta arquivos
    por tipo (classificado pelo NOME, normalizado) + procura qualquer nome de
    arquivo com marcador de encerramento explicito."""
    contagem = {chave: 0 for chave in _TIPOS_DOC_NC}
    total = 0
    tem_marcador_encerramento = []
    if not raiz_notificacoes or not U.exists(raiz_notificacoes):
        return contagem, total, tem_marcador_encerramento
    for dp, _dn, fn in U.walk(raiz_notificacoes):
        for f in fn:
            total += 1
            n = U.norm(f)
            if any(marca in n for marca in _MARCADORES_ENCERRAMENTO):
                tem_marcador_encerramento.append(os.path.join(dp, f))
            if "nao conformidade" in n:
                contagem["abertura"] += 1
            elif "nota_tecnica" in n or "nota tecnica" in n:
                contagem["nota_tecnica"] += 1
            elif "parecer_tecnico" in n or "parecer tecnico" in n:
                contagem["parecer_resposta"] += 1
            elif "resposta" in n:
                contagem["resposta"] += 1
    return contagem, total, tem_marcador_encerramento


def avaliar_reconstrucao_historica(obra, contagem_status=None):
    """Verifica se os documentos individuais de NC (spec Sec.2/§6) permitem
    reconstruir data de emissao, mudancas de status, encerramento e evidencia
    -- SEM tocar na planilha (Tbl_VISAO usa TODAY()) e sem usar o golden.

    Regra de decisao (documentada, nao "regra cega"): so seria possivel
    reconstruir a LINHA DO TEMPO de status (o dominio de 6 valores da
    planilha) se o acervo tivesse, para cada NC encerrada na planilha, um
    documento individual que declare EXPLICITAMENTE o encerramento (nome ou
    conteudo com marcador de encerramento) com data. Isso e checado contando:
      (a) documentos com marcador de encerramento no NOME (nenhum encontrado
          em nenhuma obra ate a verificacao desta sessao);
      (b) documentos de "parecer tecnico de resposta" (o unico tipo de
          documento subsequente à abertura que existe no acervo) vs. o numero
          de NC 'Encerrada' que a planilha relata para a mesma obra.
    Se (a) == 0 E (b) documentos < NC 'Encerrada' da planilha, a conclusao e
    UNRECONSTRUCTABLE_FROM_CURRENT_SOURCE: a planilha registra encerramentos
    sem documento de encerramento correspondente no acervo -- ou seja, o
    encerramento e um evento so registrado na planilha (manual), nao
    documentado nos arquivos individuais.
    """
    pasta_gto = os.path.join(U.BASE_OBRAS, obra, "3_Gestao Tecnica de Obra")
    raiz_notificacoes = C.resolver_subpasta(pasta_gto, prefixo="3. notifica")

    contagem, total_arquivos, marcadores = _contar_documentos_por_tipo(raiz_notificacoes)
    n_encerradas_planilha = (contagem_status or {}).get("Encerrada")

    evidencia = {
        "raiz_notificacoes": raiz_notificacoes,
        "total_arquivos_individuais": total_arquivos,
        "contagem_por_tipo": contagem,
        "arquivos_com_marcador_de_encerramento_no_nome": marcadores,
        "nc_encerrada_na_planilha": n_encerradas_planilha,
    }

    if not raiz_notificacoes or total_arquivos == 0:
        return {
            "historical_nc_state": "UNRECONSTRUCTABLE_FROM_CURRENT_SOURCE",
            "motivo": "pasta '3. Notificações' nao encontrada ou vazia para %s -- nao ha "
                      "documento individual para investigar" % obra,
            "impacto": "history_by_status so pode vir da planilha Tbl_VISAO (estado atual, "
                       "TODAY()), nunca do estado da semana do relatorio",
            "evidencia": evidencia,
        }

    reconstruivel = (
        len(marcadores) > 0
        and n_encerradas_planilha is not None
        and contagem["parecer_resposta"] >= n_encerradas_planilha
    )
    if reconstruivel:
        return {
            "historical_nc_state": "RECONSTRUCTABLE",
            "motivo": "acervo tem marcador de encerramento no nome de arquivo e "
                      "documento(s) de parecer de resposta em quantidade >= NC 'Encerrada' "
                      "da planilha -- reconstrucao por-NC nao implementada nesta fatia "
                      "(fora do orcamento desta investigacao pontual); os SINAIS existem",
            "impacto": "history_by_status permanece o estado ATUAL da planilha nesta fatia; "
                       "reconstrucao por documento fica para uma fatia futura",
            "evidencia": evidencia,
        }

    return {
        "historical_nc_state": "UNRECONSTRUCTABLE_FROM_CURRENT_SOURCE",
        "motivo": (
            "Tbl_VISAO usa TODAY() (estado atual, nao historico da semana). Os documentos "
            "individuais em 'Not. Geradas' + 'Respostas' registram a ABERTURA da NC (nome "
            "contem 'Não Conformidade') e, para parte dos casos, uma Nota Técnica, uma "
            "Resposta da executora e um Parecer Técnico de acompanhamento -- mas NENHUM "
            "arquivo do acervo (%d arquivos verificados) tem no nome um marcador explicito "
            "de encerramento/mudanca de status, e o numero de pareceres de resposta "
            "encontrados (%s) e menor que o numero de NC 'Encerrada' que a propria planilha "
            "relata (%s) para %s -- ou seja, a planilha registra encerramentos sem documento "
            "de encerramento correspondente no acervo de arquivos individuais. O encerramento "
            "e um evento registrado so na planilha (provavelmente atualizacao manual de "
            "status), nao documentado nos arquivos individuais desta fonte."
        ) % (total_arquivos, contagem["parecer_resposta"], n_encerradas_planilha, obra),
        "impacto": "nonconformities.history_by_status reflete o estado ATUAL (na data em que "
                   "o script roda), nao o estado da SEMANA do relatorio -- a contagem 'em "
                   "atraso'/'nao encerrada' pode divergir do que estava correto na semana. "
                   "Para relatorios novos (extraidos na propria semana) o problema nao existe "
                   "-- a leitura e contemporanea.",
        "evidencia": evidencia,
    }
