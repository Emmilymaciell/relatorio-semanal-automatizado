# -*- coding: utf-8 -*-
r"""
scripts/gerador/source_discovery.py - Slice 1: descobre ONDE mora cada fonte
para uma obra + periodo, e produz o `manifest.yaml` (sem ler dado de negocio
das fontes -- so metadados de arquivo: existe, tamanho, sha1, mtime).

Layout ja verificado (spec Sec.2 + verificacao direta desta sessao); a
resolucao de nomes variaveis de pasta usa config/source_layout.yaml via
`gerador.config.resolver_subpasta/resolver_arquivo` (prefixo ou igualdade
normalizados) -- nunca `if obra == "..."`.
"""
import os
import re
import sys
import datetime as dt

AQUI = os.path.dirname(os.path.abspath(__file__))
SCRIPTS = os.path.dirname(AQUI)
if SCRIPTS not in sys.path:
    sys.path.insert(0, SCRIPTS)
import prog_util as U  # noqa: E402

from . import config as C  # noqa: E402
from . import periodo as P  # noqa: E402
from .extratores import nc as EX_NC  # noqa: E402
from .extratores import ambiental as EX_AMB  # noqa: E402

_RE_SEMANA = re.compile(
    r"^(?P<a>\d{4}\.\d{2}\.\d{2})-(?P<b>\d{4}\.\d{2}\.\d{2})\s+(?P<num>\d+)\s*[ºo°]?\s*Relat",
    re.IGNORECASE)

_LIMITE_SHA1_GRANDE = 20 * 1024 * 1024   # acima disso, sha1 parcial (1MB)


def _rx(layout, *caminho, raiz="arvore_por_obra"):
    """Le `regex` de config/source_layout.yaml (padrao tolerante de resolucao de
    pasta/arquivo, aplicado ao nome normalizado). None se nao declarado."""
    no = layout.get(raiz) or {}
    for k in caminho:
        no = (no or {}).get(k) or {}
    return no.get("regex") if isinstance(no, dict) else None


def _info_arquivo(caminho):
    """Metadados de UM arquivo (nao le conteudo). Nunca levanta excecao para
    fonte ausente/ilegivel -- devolve status."""
    if not caminho:
        return {"path": None, "exists": False, "status": "NOT_FOUND"}
    if not U.exists(caminho):
        return {"path": caminho, "exists": False, "status": "NOT_FOUND"}
    try:
        st = U.stat_info(caminho)
        limite = None if st["tamanho_bytes"] <= _LIMITE_SHA1_GRANDE else 1024 * 1024
        sha1 = U.sha1(caminho, limite)
        d = {"path": caminho, "exists": True, "status": "OK",
             "size": st["tamanho_bytes"], "mtime": st["modificado"], "sha1": sha1}
        if limite:
            d["sha1_parcial"] = True
        return d
    except OSError as e:
        return {"path": caminho, "exists": True, "status": "UNREADABLE", "erro": str(e)}


def localizar_pasta_semana(obra, data_inicio, data_fim):
    """Acha a pasta-semana cujo intervalo (min,max) de datas no nome bate
    exatamente com [data_inicio, data_fim]. Tolera pastas com as duas datas
    invertidas no nome (usa min/max, nao posicao) e ignora 'obsoleto' e
    arquivos soltos na raiz de '5. Relat*' (ex.: planilhas fora de pasta)."""
    raiz = U.pasta_relatorios(obra)
    if not raiz:
        return None, None
    alvo_lo, alvo_hi = min(data_inicio, data_fim), max(data_inicio, data_fim)
    for nome in U.listdir(raiz):
        caminho = os.path.join(raiz, nome)
        if not U.isdir(caminho) or "obsoleto" in U.norm(nome):
            continue
        m = _RE_SEMANA.match(nome)
        if not m:
            continue
        try:
            da = dt.date.fromisoformat(m.group("a").replace(".", "-"))
            db = dt.date.fromisoformat(m.group("b").replace(".", "-"))
        except ValueError:
            continue
        lo, hi = min(da, db), max(da, db)
        if lo == alvo_lo and hi == alvo_hi:
            return caminho, int(m.group("num"))
    return None, None


def construir_manifest(obra, data_inicio, data_fim):
    if isinstance(data_inicio, str):
        data_inicio = dt.date.fromisoformat(data_inicio)
    if isinstance(data_fim, str):
        data_fim = dt.date.fromisoformat(data_fim)

    info = C.obra_info(obra)
    rodovia = info["rodovia"]
    layout = C.source_layout()

    warnings, not_found = [], []

    def aviso(msg):
        warnings.append(msg)

    def marcar(chave, entrada):
        if entrada.get("status") not in ("OK",) and entrada.get("status") != "NOT_APPLICABLE":
            not_found.append({"fonte": chave, "status": entrada.get("status"),
                              "path": entrada.get("path"), "erro": entrada.get("erro")})
        return entrada

    numero = P.numero_relatorio(obra, data_inicio)
    schema_version = P.schema_version(data_inicio)

    sources = {}

    # --------------------------------------------------- pasta da semana
    pasta_semana, numero_na_pasta = localizar_pasta_semana(obra, data_inicio, data_fim)
    if not pasta_semana:
        aviso("pasta-semana nao encontrada para %s %s..%s" % (obra, data_inicio, data_fim))

    pasta_gto = os.path.join(U.BASE_OBRAS, obra, "3_Gestao Tecnica de Obra")
    pasta_semanal_raiz = U.pasta_semanal_raiz(obra)

    # --------------------------------------------------- fichas de campo diarias
    field_records = []
    missing_days = []
    if pasta_semanal_raiz:
        pasta_fichas = C.resolver_subpasta(pasta_semanal_raiz, prefixo="1. ficha")
    else:
        pasta_fichas = None
        aviso("pasta_semanal_raiz nao encontrada (prefixo '4. Relat*') para %s" % obra)

    for dia in P.dias_do_periodo(data_inicio):
        diso = dia["date"]
        caminho_dia = None
        if pasta_fichas:
            for nome in U.listdir(pasta_fichas):
                if not nome.startswith(diso):
                    continue
                n = U.norm(nome)
                if "ficha" in n and "campo" in n and os.path.splitext(nome)[1].lower() in (
                        ".xlsx", ".xlsm"):
                    caminho_dia = os.path.join(pasta_fichas, nome)
                    break
        entrada = dict(date=diso, **_info_arquivo(caminho_dia))
        field_records.append(entrada)
        if entrada["status"] == "NOT_FOUND":
            missing_days.append(diso)
    sources["field_records"] = field_records
    sources["field_records_missing_days"] = missing_days

    # --------------------------------------------------- pasta de fotos
    pasta_fotos = C.resolver_subpasta(pasta_semanal_raiz, iguais=["3. fotos"]) \
        if pasta_semanal_raiz else None
    photos = []
    for dia in P.dias_do_periodo(data_inicio):
        diso = dia["date"]
        d = dt.date.fromisoformat(diso)
        alvo_h = "%02d-%02d-%02d" % (d.year % 100, d.month, d.day)
        alvo_u = alvo_h.replace("-", "_")
        pasta_dia = None
        if pasta_fotos and U.exists(pasta_fotos):
            for nome in U.listdir(pasta_fotos):
                if nome in (alvo_h, alvo_u):
                    pasta_dia = os.path.join(pasta_fotos, nome)
                    break
        if not pasta_dia:
            photos.append({"date": diso, "status": "NOT_FOUND"})
            continue
        for servico in U.listdir(pasta_dia):
            sp = os.path.join(pasta_dia, servico)
            if not U.isdir(sp):
                continue
            arquivos = [f for f in U.listdir(sp) if not U.isdir(os.path.join(sp, f))]
            photos.append({"date": diso, "service_folder": servico,
                           "files": arquivos, "n": len(arquivos), "status": "OK"})
    sources["photos"] = photos

    # --------------------------------------------------- auxiliares da semana
    pasta_aux = C.resolver_subpasta(pasta_semana, iguais=["arquivos auxiliares"]) \
        if pasta_semana else None
    if pasta_semana and not pasta_aux:
        aviso("'Arquivos Auxiliares' nao encontrada dentro da pasta-semana de %s" % obra)

    acumulada = C.resolver_arquivo(pasta_aux, prefixo="ficha de campo acumulada",
                                    extensoes=[".xlsx"],
                                    regex=_rx(layout, "ficha_acumulada", raiz="auxiliares")) if pasta_aux else None
    sources["accumulated_sheet"] = marcar("accumulated_sheet", _info_arquivo(acumulada))

    grafico = C.resolver_arquivo(pasta_aux, iguais=["grafico - avanco global"],
                                  extensoes=[".xlsx"]) if pasta_aux else None
    sources["progress_chart"] = marcar("progress_chart", _info_arquivo(grafico))

    prog_semana = C.resolver_arquivo(pasta_aux, prefixo="[programa] avanco global de obra",
                                      extensoes=[".xlsx"]) if pasta_aux else None
    sources["prog_progress"] = marcar("prog_progress", _info_arquivo(prog_semana))

    relfoto = C.resolver_arquivo(pasta_aux, prefixo="relatorio fotografico",
                                  extensoes=[".xlsx"]) if pasta_aux else None
    sources["photo_workbook"] = marcar("photo_workbook", _info_arquivo(relfoto))

    # --------------------------------------------------- fontes corporativas / compartilhadas
    compartilhado = os.path.join(U.SHAREPOINT, "Programa - 2. Compartilhado interno",
                                 "03_Gestao dos contratos")
    unifilar_compartilhado = os.path.join(
        compartilhado, "UNIFILAR (com não conformidades)", "UNIFILAR %s.xlsx" % info["rodovia_sem_hifen"][2:])
    sources["progress_unifilar"] = marcar("progress_unifilar", _info_arquivo(unifilar_compartilhado))

    nc_tracker_path = os.path.join(U.BASE_OBRAS, "02_Acompanhamento de NC automatizada.xlsm")
    sources["nc_tracker"] = marcar("nc_tracker", _info_arquivo(nc_tracker_path))

    # Caminho vem de config/source_layout.yaml (corporativas.licenca_ambiental),
    # nunca hardcoded aqui -- decisao de qual copia e' a operacional (raiz
    # vs. copias do relatorio mensal) investigada e registrada em
    # docs/fontes/acompanhamento_ambiental.md Sec.6 (Fase 3.4, SOURCE FIX).
    licenca_ambiental_path = C.corporativa("licenca_ambiental")

    # --------------------------------------------------- NC emitidas no periodo (para localizar PDFs)
    nc_documents = []
    if sources["nc_tracker"]["status"] == "OK":
        nc_info = EX_NC.ler(nc_tracker_path, rodovia, data_inicio, data_fim)
        if nc_info.get("status") == "OK":
            idx_pdfs = EX_NC.indexar_pdfs(obra)
            for item in nc_info["issued_in_period"]:
                doc = item.get("documento")
                pdfs = EX_NC.pdfs_da_nc(idx_pdfs, doc)
                nc_documents.append({"documento": doc, "data_abertura": item["data_abertura"],
                                     "pdfs": pdfs, "status": "OK" if pdfs else "NOT_FOUND"})
    sources["nc_documents"] = nc_documents

    # --------------------------------------------------- ambiental
    pasta_amb_sst = C.resolver_subpasta(pasta_gto, prefixo="6. acompanhamento ambiental")
    pasta_visita = EX_AMB.localizar_pasta_visita(pasta_amb_sst, data_inicio, data_fim) \
        if pasta_amb_sst else None
    fichas_visita = EX_AMB.ler_fichas_da_visita(pasta_visita) if pasta_visita else \
        {"status": "NOT_APPLICABLE"}
    applies = EX_AMB.ciclo_aplica(data_inicio)
    sources["environmental"] = {
        "status": "OK" if fichas_visita.get("status") == "OK" else fichas_visita.get("status"),
        "applies": applies,
        "licenca_ambiental_planilha": marcar("environmental.planilha", _info_arquivo(licenca_ambiental_path)),
        "pasta_visita": pasta_visita,
        "fa_ambiental": marcar("environmental.fa_ambiental",
                               _info_arquivo(fichas_visita.get("fa_ambiental"))),
        "fa_sst": marcar("environmental.fa_sst", _info_arquivo(fichas_visita.get("fa_sst"))),
    }
    if applies and fichas_visita.get("status") != "OK":
        aviso("ciclo ambiental quinzenal aplica para %s %s, mas pasta de visita nao "
              "foi encontrada dentro do periodo" % (obra, data_inicio))

    # --------------------------------------------------- diario de obras
    pasta_diario = C.resolver_subpasta(pasta_semanal_raiz, prefixo="4. di") \
        if pasta_semanal_raiz else None
    diario_path = None
    if pasta_diario:
        pasta_assinado = C.resolver_subpasta(pasta_diario, prefixo="diario de obra digital assinado",
                                              regex=_rx(layout, "pasta_diario_obras", "subpastas", "assinado"))
        alvo_txt = "%s a %s" % (data_inicio.isoformat(), data_fim.isoformat())
        if pasta_assinado:
            for nome in U.listdir(pasta_assinado):
                if nome.startswith(alvo_txt) and nome.lower().endswith(".pdf"):
                    diario_path = os.path.join(pasta_assinado, nome)
                    break
    sources["daily_reports"] = marcar("daily_reports", _info_arquivo(diario_path))

    # --------------------------------------------------- reunioes (ata) no periodo
    pasta_ata = C.resolver_subpasta(pasta_gto, prefixo="9. ata de reuni",
                                  regex=_rx(layout, "pasta_ata_reuniao"))
    meetings = []
    if pasta_ata and U.exists(pasta_ata):
        for nome in U.listdir(pasta_ata):
            m = re.search(r"(\d{4}-\d{2}-\d{2})", nome)
            if not m:
                continue
            try:
                d = dt.date.fromisoformat(m.group(1))
            except ValueError:
                continue
            if data_inicio <= d <= data_fim:
                meetings.append(dict(date=d.isoformat(),
                                     **_info_arquivo(os.path.join(pasta_ata, nome))))
    sources["meetings"] = meetings

    # --------------------------------------------------- ficha acompanhamento semanal (so registrar)
    pasta_ficha_acomp = C.resolver_subpasta(pasta_semanal_raiz, prefixo="2. ficha") \
        if pasta_semanal_raiz else None
    sources["weekly_tracking_sheet_folder"] = {
        "path": pasta_ficha_acomp, "exists": bool(pasta_ficha_acomp),
        "nota": "NAO alimenta o v3 (P-08); apenas registrado"}

    # --------------------------------------------------- anexos candidatos (existencia bruta)
    annexes = []
    for chave, caminho in (("progress_unifilar (diagrama)", unifilar_compartilhado),
                           ("daily_reports (diario)", diario_path)):
        annexes.append({"fonte": chave, "path": caminho, "exists": bool(caminho and U.exists(caminho))})
    sources["annexes"] = annexes

    # --------------------------------------------------- golden (SO metadados)
    # NOTA: source_discovery e o UNICO modulo autorizado a REGISTRAR o caminho
    # do golden (spec Sec.5) -- por isso NAO chama assert_nao_e_golden aqui (so
    # grava caminho/sha1/n_paginas, nunca abre o arquivo para extrair dado de
    # negocio). Todo EXTRATOR chama a guarda antes de abrir um arquivo; e
    # tests/test_sem_vazamento_golden.py prova que ela dispara para estes
    # mesmos caminhos.
    golden = {"docx": {"path": None, "sha1": None},
             "pdf": {"path": None, "sha1": None, "n_paginas": None},
             "aviso": "REFERENCIA DE COMPARACAO -- PROIBIDO COMO FONTE"}
    if pasta_semana:
        for nome in U.listdir(pasta_semana):
            caminho = os.path.join(pasta_semana, nome)
            if U.isdir(caminho):
                continue
            if re.match(r"^RO_\d{3}\.\d+\s*-\s*.+\.docx$", nome, re.IGNORECASE):
                golden["docx"] = {"path": caminho, "sha1": U.sha1(caminho)}
            elif re.match(r"^RO_\d{3}\.\d+\s*-\s*.+\.pdf$", nome, re.IGNORECASE):
                n_paginas = None
                try:
                    import fitz
                    local = U.copiar_para_tmp(caminho, "golden_meta", "golden.pdf")
                    d = fitz.open(local)
                    n_paginas = d.page_count
                    d.close()
                except Exception as e:
                    aviso("nao foi possivel contar paginas do golden PDF: %s" % e)
                golden["pdf"] = {"path": caminho, "sha1": U.sha1(caminho), "n_paginas": n_paginas}

    manifest = {
        "report": {
            "road": rodovia, "road_key": obra,
            "period_start": data_inicio.isoformat(), "period_end": data_fim.isoformat(),
            "report_number": numero, "revision": 1, "schema_version": schema_version,
            "report_number_na_pasta_semana": numero_na_pasta,
        },
        "sources": sources,
        "golden_reference": golden,
        "discovery": {
            "resolved_at": dt.datetime.now().astimezone().isoformat(timespec="seconds"),
            "warnings": warnings,
            "not_found": not_found,
        },
    }
    return manifest
