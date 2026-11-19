# -*- coding: utf-8 -*-
r"""
scripts/gerador/canonical.py - Slice 2: modelo canonico (`report_data.json`) +
proveniencia paralela (`report_provenance.json`), a partir do manifest do
Slice 1 (source_discovery.construir_manifest).

Chaves de topo (18, config/report_schema.yaml + spec Sec.3):
    metadata, road, contract, period, report_number, services, climate,
    progress, field_records, photos, nonconformities, environment, sst,
    meetings, risks, interferences, annexes, closing_term

Nunca inventa numero/data/estaca/NC/percentual: todo campo sem fonte segura
fica None com um `_status` irmao (OK | NOT_FOUND | UNREADABLE |
NOT_APPLICABLE | MVP_LIMITATION), e some's para `discovery.warnings`.
"""
import os
import sys
import datetime as dt

AQUI = os.path.dirname(os.path.abspath(__file__))
SCRIPTS = os.path.dirname(AQUI)
if SCRIPTS not in sys.path:
    sys.path.insert(0, SCRIPTS)
import prog_util as U  # noqa: E402

from . import config as C  # noqa: E402
from . import periodo as P  # noqa: E402
from .provenance import Provenance  # noqa: E402
from .source_discovery import construir_manifest  # noqa: E402
from .extratores import ficha_campo, ficha_acumulada, avanco, nc, ambiental, fotos, anexos, parecer  # noqa: E402

_CLASSE_POR_ALIAS_CACHE = {}


def _mapa_abas_para_taxonomia():
    tax = C.services_taxonomy()
    return tax.get("abas_relatorio_fotografico", {})


def _canonico_por_codigo_ficha(obra, codigo):
    """Acha a chave canonica (ex.: 'TERRAPLENAGEM') cujo codigo_ficha (ou
    codigo_ficha_por_obra[obra]) bate com `codigo`."""
    tax = C.services_taxonomy()
    for grupo in ("servicos", "controle_tecnologico", "topografia",
                  "meio_ambiente_e_sst", "eventos"):
        for chave, ent in tax.get(grupo, {}).items():
            cpo = ent.get("codigo_ficha_por_obra")
            if cpo and cpo.get(obra) == codigo:
                return chave, ent
            cod = ent.get("codigo_ficha")
            if cod is None:
                continue
            codes = cod if isinstance(cod, list) else [cod]
            if codigo in codes:
                return chave, ent
    return None, None


def _canonical_name(chave, ent):
    return ent.get("canonical_name") if ent else chave


def _entrada_clima(diso, label, clima, status, motivo=None):
    """Monta UMA entrada de climate.by_day (Q-8, spec Sec.1).

    Quando `clima` nao foi comprovado por nenhuma fonte (ficha ausente,
    ilegivel, ou sem marca de X legivel), a entrada NAO fica muda: ganha
    `clima_label: "INDETERMINADO"` e `validation_status: "HITL"` -- nunca um
    default ("Bom") nem repeticao do dia anterior. `clima`/`_status` continuam
    descrevendo o resultado da EXTRACAO (None/NOT_FOUND etc, como antes);
    `validation_status` e `clima_label` descrevem se o dado esta pronto para
    o documento final -- sao conceitos distintos e coexistem de proposito."""
    entrada = {"date": diso, "label": label, "clima": clima, "_status": status}
    if clima is None:
        entrada["clima_label"] = "INDETERMINADO"
        entrada["validation_status"] = "HITL"
        entrada["motivo"] = motivo or "sem Ficha de Campo e sem fonte alternativa comprovada"
    else:
        entrada["clima_label"] = clima
        entrada["validation_status"] = "OK"
    return entrada


def construir(obra, data_inicio, data_fim):
    if isinstance(data_inicio, str):
        data_inicio = dt.date.fromisoformat(data_inicio)
    if isinstance(data_fim, str):
        data_fim = dt.date.fromisoformat(data_fim)

    manifest = construir_manifest(obra, data_inicio, data_fim)
    prov = Provenance()
    info = C.obra_info(obra)
    dias = P.dias_do_periodo(data_inicio)
    warnings = list(manifest["discovery"]["warnings"])

    # ============================================================ metadata
    metadata = {
        "generated_at": dt.datetime.now().astimezone().isoformat(timespec="seconds"),
        "generator": "scripts/shadow_report.py (Fase 3, Slice 1+2)",
        "schema_version": manifest["report"]["schema_version"],
        "road_key": obra, "road": info["rodovia"],
        "aviso": "RELATORIO-SOMBRA -- gerado em paralelo para validacao; nao substitui o oficial",
    }

    # ============================================================ road / contract
    road = {k: info.get(k) for k in (
        "rodovia", "trecho", "trecho_extenso", "extensao_km", "extensao_texto",
        "codigo_sre", "lote", "empresa_executora_capa", "empresa_executora_contrato",
        "empresa_supervisora")}
    contract = {}
    # §7 (spec Sec.3): todo campo contratual do gerador vem do cadastro mestre
    # (road_registry.yaml), NUNCA do golden. `contrato_provenance` (por obra)
    # registra arquivo + trecho de origem de cada campo -- espelhado aqui em
    # report_provenance.json para que o DOCX (fora de escopo) tenha de onde puxar.
    contrato_prov_obra = info.get("contrato_provenance", {})
    for k in ("contrato_obra", "contrato_prog", "instrumento_convocacao",
              "contrato_assinado_em", "ordem_servico", "ordem_servico_em"):
        v = info.get(k)
        if v is None:
            v = C.road_registry().get("meta", {}).get(k)
        contract[k] = v
        origem = contrato_prov_obra.get(k)
        if v is None:
            contract["_status_%s" % k] = "NOT_FOUND"
            motivo = (origem or {}).get("motivo", "campo nao comprovado por documento acessivel "
                                                   "a esta fatia; nao copiado do golden")
            warnings.append("contract.%s %s: NOT_FOUND -- %s" % (k, obra, motivo))
            prov.registrar("contract.%s" % k, None, method="direct extraction (busca pontual)",
                           confidence="NOT_FOUND", status="NOT_FOUND", validation_status="HITL",
                           nota=motivo)
        elif origem:
            prov.registrar("contract.%s" % k, v, source_file=origem.get("arquivo"),
                           method="direct extraction (busca pontual, PDF do contrato/OS)",
                           confidence="CONFIRMADO POR EVIDÊNCIA DIRETA",
                           cell=origem.get("pagina"), nota=origem.get("trecho"))
        else:
            # campo herdado sem `contrato_provenance` dedicada (ex.: RD-101,
            # confirmado na capa/termo de encerramento na fase de mapeamento)
            prov.registrar("contract.%s" % k, v, source_file="config/road_registry.yaml",
                           method="cadastro mestre (mapeamento fase 2)",
                           confidence="CONFIRMADO")

    # ============================================================ period / report_number
    period = {"start": data_inicio.isoformat(), "end": data_fim.isoformat(), "days": dias}
    report_number = manifest["report"]["report_number"]
    prov.registrar("report_number", report_number,
                   source_file="config/road_registry.yaml", method="derived",
                   confidence="CONFIRMADO",
                   inputs=["road_registry.obras.%s.ancora_semana_1" % obra, "period.start"])

    # ============================================================ field_records + clima
    field_records = []
    climate_by_day = []
    codigos_vistos = set()

    for dia in dias:
        diso = dia["date"]
        fonte = next((f for f in manifest["sources"]["field_records"] if f["date"] == diso), None)
        if not fonte or fonte.get("status") != "OK":
            status_fonte = fonte.get("status") if fonte else "NOT_FOUND"
            field_records.append({"date": diso, "status": status_fonte, "servicos": None})
            motivo = ("sem Ficha de Campo e sem fonte alternativa comprovada"
                      if status_fonte == "NOT_FOUND" else
                      "ficha de campo com problema de leitura (%s) e sem fonte alternativa "
                      "comprovada" % status_fonte)
            climate_by_day.append(_entrada_clima(diso, dia["label"], None, status_fonte, motivo))
            prov.registrar("climate.by_day.%s" % diso, None,
                           source_file=(fonte or {}).get("path"),
                           method="direct extraction", confidence="NOT_FOUND",
                           status=status_fonte, validation_status="HITL",
                           nota=motivo)
            continue

        leitura = ficha_campo.ler(fonte["path"])
        if leitura["status"] != "OK":
            field_records.append({"date": diso, "status": leitura["status"], "servicos": None,
                                  "erro": leitura.get("erro")})
            motivo = ("ficha de campo existe mas nao pode ser lida (%s: %s) -- sem fonte "
                      "alternativa comprovada" % (leitura["status"], leitura.get("erro", "")))
            climate_by_day.append(_entrada_clima(diso, dia["label"], None, leitura["status"], motivo))
            warnings.append("ficha_campo %s %s: %s" % (obra, diso, leitura["status"]))
            prov.registrar("climate.by_day.%s" % diso, None, source_file=fonte["path"],
                           method="direct extraction", confidence="NOT_FOUND",
                           status=leitura["status"], validation_status="HITL", nota=motivo)
            continue

        field_records.append({"date": diso, "status": "OK", "servicos": leitura["servicos"]})
        for linha in leitura["servicos"]:
            if linha.get("servico_codigo"):
                codigos_vistos.add(linha["servico_codigo"])

        clima = leitura["clima"]
        status_clima = "OK" if clima else leitura["clima_evidencia"]["status"]
        motivo_clima = None if clima else (
            "Ficha de Campo do dia existe mas a linha de marca do clima (X) nao foi lida com "
            "seguranca (%s): %s" % (leitura["clima_evidencia"]["status"],
                                    leitura["clima_evidencia"].get("motivo", "")))
        climate_by_day.append(_entrada_clima(diso, dia["label"], clima, status_clima, motivo_clima))
        chave_prov = "climate.by_day.%s" % diso
        if clima:
            prov.registrar(chave_prov, clima, source_file=fonte["path"], sheet=ficha_campo.ABA,
                           cell=leitura["clima_evidencia"].get("celula"), validation_status="OK")
        else:
            prov.registrar(chave_prov, None, source_file=fonte["path"], sheet=ficha_campo.ABA,
                           method="direct extraction", confidence="NOT_FOUND",
                           validation_status="HITL", nota=motivo_clima,
                           **{"status": leitura["clima_evidencia"]["status"]})

    days_without_work = list(manifest["sources"]["field_records_missing_days"])
    climate = {
        "by_day": climate_by_day,
        "days_without_work": days_without_work,
        "frases_sem_execucao": ["No dia %s não houve execução de atividades."
                                % dt.date.fromisoformat(d).strftime("%d/%m/%Y")
                                for d in days_without_work],
    }
    prov.derivado("climate.days_without_work", days_without_work,
                  inputs=["sources.field_records[*].status"])

    # ============================================================ services
    mapa_abas = _mapa_abas_para_taxonomia()
    relfoto_info = manifest["sources"]["photo_workbook"]
    servicos_chaves = []
    ev_fotos = {"status": "NOT_FOUND"}
    if relfoto_info.get("status") == "OK":
        ev_fotos = fotos.abas_com_foto(relfoto_info["path"])
        if ev_fotos.get("status") == "OK":
            for aba_nome in ev_fotos["subsecoes_5_2"]:
                destino = mapa_abas.get(aba_nome)
                if isinstance(destino, dict):
                    continue  # aba explicitamente sem classe (CAPA, AVANÇO_GERAL)
                if destino:
                    servicos_chaves.append(destino)
                else:
                    warnings.append("services: aba '%s' do relatorio fotografico sem "
                                    "mapeamento em services_taxonomy.abas_relatorio_fotografico"
                                    % aba_nome)
        else:
            warnings.append("services: nao foi possivel derivar do relatorio fotografico (%s)"
                            % ev_fotos.get("status"))
    else:
        warnings.append("services: photo_workbook nao encontrado (%s)" % relfoto_info.get("status"))

    tax = C.services_taxonomy()
    ordem = tax.get("ordem_subsecoes_5_2", [])
    servicos_ordenados = [c for c in ordem if c in servicos_chaves] + \
        [c for c in servicos_chaves if c not in ordem]
    services = []
    for chave in servicos_ordenados:
        ent = None
        # Q-9: procurar em TODOS os grupos da taxonomia (nao so "servicos") --
        # abas do relatorio fotografico podem mapear para chaves de outros
        # grupos (ex.: ACOMPANHAMENTO_AMBIENTAL em meio_ambiente_e_sst,
        # REGISTRO_ADMINISTRATIVO em eventos); restringir a "servicos" fazia
        # o canonical_name cair no fallback (a propria chave, ex.
        # "ACOMPANHAMENTO_AMBIENTAL") em vez do nome redigido da taxonomia.
        for grupo in ("servicos", "controle_tecnologico", "topografia",
                      "meio_ambiente_e_sst", "eventos"):
            ent = tax.get(grupo, {}).get(chave)
            if ent:
                break
        services.append({
            "key": chave,
            "canonical_name": _canonical_name(chave, ent),
            # Q-9: sigla vem SO de config/services_taxonomy.yaml (campo
            # `sigla`), nunca de heuristica sobre o texto. None quando a
            # taxonomia nao registra sigla para a entrada -- o redator usa so
            # o nome canonico (regra de 1a ocorrencia e do content builder,
            # fora do escopo desta fatia).
            "sigla": ent.get("sigla") if ent else None,
        })
    prov.registrar("services", [s["canonical_name"] for s in services],
                   source_file=relfoto_info.get("path"), method="derived",
                   confidence="CONFIRMADO POR CRUZAMENTO DE FONTES (regra T-17)",
                   inputs=["abas visible+com foto do RELATORIO FOTOGRAFICO"])

    # ============================================================ progress
    progress = {}
    prog_info = manifest["sources"]["prog_progress"]
    if prog_info.get("status") == "OK":
        leitura = avanco.ler_prog(prog_info["path"], obra)
        if leitura["status"] == "OK":
            progress["global_concluido"] = leitura["c3"]
            progress["global_concluido_atacado"] = leitura["c7"]
            prov.registrar("progress.global_concluido", leitura["c3"], source_file=prog_info["path"],
                           sheet=leitura["aba"], cell="C3")
            prov.registrar("progress.global_concluido_atacado", leitura["c7"],
                           source_file=prog_info["path"], sheet=leitura["aba"], cell="C7")
        else:
            progress["global_concluido"] = None
            progress["global_concluido_atacado"] = None
            progress["_status"] = leitura["status"]
            warnings.append("progress (Programa C3/C7) %s: %s" % (obra, leitura["status"]))
    else:
        progress["global_concluido"] = None
        progress["global_concluido_atacado"] = None
        progress["_status"] = prog_info.get("status")

    chart_info = manifest["sources"]["progress_chart"]
    if chart_info.get("status") == "OK":
        leitura_g = avanco.ler_grafico(chart_info["path"])
        progress["series"] = leitura_g.get("series")
    else:
        progress["series"] = None
        progress["_status_series"] = chart_info.get("status")

    # ============================================================ photos
    photos_by_day = {}
    for item in manifest["sources"]["photos"]:
        photos_by_day.setdefault(item["date"], []).append(item)
    photos_out = []
    for dia in dias:
        itens = photos_by_day.get(dia["date"], [])
        if len(itens) == 1 and itens[0].get("status") == "NOT_FOUND":
            photos_out.append({"date": dia["date"], "status": "NOT_FOUND", "services": []})
            continue
        servicos_dia = [{"service_folder": it["service_folder"], "n": it["n"]}
                        for it in itens if it.get("status") == "OK"]
        photos_out.append({"date": dia["date"], "status": "OK", "services": servicos_dia,
                           "total": sum(s["n"] for s in servicos_dia)})
    photos = {
        "by_day": photos_out,
        "overlay_read": False,
        "status": "READY_WITH_HITL",
        "nota": "inventario de arquivos apenas; sem OCR do overlay (data/UTM/estaca/servico "
                "impressos na foto) nesta iteracao",
    }

    # ============================================================ nonconformities
    nc_info = manifest["sources"]["nc_tracker"]
    if nc_info.get("status") == "OK":
        leitura_nc = nc.ler(nc_info["path"], info["rodovia"], data_inicio, data_fim)
    else:
        leitura_nc = {"status": nc_info.get("status")}

    if leitura_nc.get("status") == "OK":
        nonconformities = {
            "issued_in_period": leitura_nc["issued_in_period"],
            "history_by_status": leitura_nc["history_by_status"],
            "history_by_status_alerta": leitura_nc["alerta"],
            "orgao": {"value": "não houve", "confidence": "DESCONHECIDO_SE_HOUVER",
                       "nota": "133/133 casos observados na amostra; comportamento com NC "
                               "da Órgão Contratante e desconhecido"},
        }
        prov.registrar("nonconformities.issued_in_period", leitura_nc["issued_in_period"],
                       source_file=nc_info["path"], sheet="VISAO_GERAL", cell="Tbl_VISAO (A3:AA156)",
                       confidence="CONFIRMADO POR EVIDÊNCIA DIRETA")
        prov.registrar("nonconformities.history_by_status", leitura_nc["history_by_status"],
                       source_file=nc_info["path"], sheet="VISAO_GERAL", cell="Tbl_VISAO (A3:AA156)",
                       confidence="CONFIRMADO POR EVIDÊNCIA DIRETA -- estado ATUAL (TODAY()), nao "
                                  "historico reconstruido da semana")
    else:
        nonconformities = {"issued_in_period": None, "history_by_status": None,
                           "_status": leitura_nc.get("status"),
                           "orgao": {"value": "não houve", "confidence": "DESCONHECIDO_SE_HOUVER"}}
        warnings.append("nonconformities %s: %s" % (obra, leitura_nc.get("status")))

    # ============================================================ environment / sst
    env_src = manifest["sources"]["environmental"]
    environment = {
        "applies": env_src["applies"],
        "cycle_anchor_date": "2026-06-06",
        "visit_date": os.path.basename(env_src["pasta_visita"]) if env_src.get("pasta_visita") else None,
        "fa_ambiental": {"path": env_src["fa_ambiental"].get("path"),
                        "exists": env_src["fa_ambiental"].get("exists", False),
                        "sha1": env_src["fa_ambiental"].get("sha1")},
        "kpis": None, "kpis_status": "MVP_LIMITATION",
        "kpis_nota": "reconstrucao do dashboard (indices/condicionantes/gestao) fora do "
                     "orcamento desta fatia -- ver scripts/12_analisar_ambiental.py",
    }
    sst = {
        "applies": env_src["applies"],
        "visit_date": environment["visit_date"],
        "fa_sst": {"path": env_src["fa_sst"].get("path"),
                  "exists": env_src["fa_sst"].get("exists", False),
                  "sha1": env_src["fa_sst"].get("sha1")},
    }
    prov.derivado("environment.applies", env_src["applies"],
                  inputs=["period.start", "ancora quinzenal 2026-06-06"],
                  confidence="CONFIRMADO POR EVIDÊNCIA DIRETA (regra A2, tests/test_regras_anexos.py)")

    # ============================================================ meetings
    meetings = manifest["sources"]["meetings"]

    # ============================================================ risks / interferences (fora do escopo provado nesta fatia)
    risks = {"value": None, "_status": "MVP_LIMITATION",
            "nota": "sem cadeia de extracao determinada nesta fatia; ver report_schema.yaml s9_interferencias"}
    interferences = {"value": None, "_status": "MVP_LIMITATION",
                     "nota": "idem risks -- campos_extraiveis mapeados na spec, mas sem "
                             "fonte por planilha comprovada; nao inventado"}

    # ============================================================ parecer tecnico (Q-11)
    if env_src["applies"]:
        parecer_resultado = parecer.avaliar(obra, data_inicio, data_fim)
        if parecer_resultado["status"] == "MATCH":
            doc = parecer_resultado["documento"]
            prov.registrar("annexes.parecer_tecnico", doc["numero"], source_file=doc["path"],
                           method=parecer_resultado["mecanismo"],
                           confidence="CONFIRMADO POR EVIDÊNCIA DIRETA (leitura do texto do PDF)",
                           evidencia=parecer_resultado.get("evidencia"))
        elif parecer_resultado["status"] == "HITL":
            prov.registrar("annexes.parecer_tecnico",
                           [c["numero"] for c in parecer_resultado["candidatos"]],
                           method=parecer_resultado["mecanismo"], confidence="HITL",
                           nota=parecer_resultado.get("motivo"))
        else:
            prov.registrar("annexes.parecer_tecnico", None,
                           method=parecer_resultado["mecanismo"], confidence="NOT_FOUND",
                           nota=parecer_resultado.get("motivo"))
    else:
        parecer_resultado = None

    # ============================================================ §6 -- estado historico das NC
    nonconformities["historical_reconstruction"] = nc.avaliar_reconstrucao_historica(
        obra, contagem_status=(leitura_nc.get("history_by_status")
                               if leitura_nc.get("status") == "OK" else None))

    # ============================================================ annexes
    nc_docs = manifest["sources"]["nc_documents"]
    annexes = anexos.montar(
        meetings=meetings, ambiental_applies=env_src["applies"],
        fa_ambiental=environment["fa_ambiental"]["path"] if environment["fa_ambiental"]["exists"] else None,
        fa_sst=sst["fa_sst"]["path"] if sst["fa_sst"]["exists"] else None,
        nc_no_periodo=nc_docs,
        unifilar_path=manifest["sources"]["progress_unifilar"].get("path"),
        unifilar_exists=manifest["sources"]["progress_unifilar"].get("exists", False),
        diario_path=manifest["sources"]["daily_reports"].get("path"),
        diario_exists=manifest["sources"]["daily_reports"].get("exists", False),
        parecer=parecer_resultado,
    )
    prov.derivado("annexes", [a["id"] for a in annexes],
                  inputs=["report_schema.yaml s11_anexos", "manifest.sources"])

    # ============================================================ ficha acumulada (so enriquece o manifest; nenhum campo do report_data depende dela nesta fatia)
    acc_info = manifest["sources"]["accumulated_sheet"]
    if acc_info.get("status") == "OK":
        leitura_acc = ficha_acumulada.ler(acc_info["path"])
        acc_info.update({k: v for k, v in leitura_acc.items() if k not in ("status", "path")})

    # ============================================================ closing_term
    closing_term = None  # placeholder explicito -- depende da paginacao final (fora de escopo)

    report_data = {
        "metadata": metadata, "road": road, "contract": contract, "period": period,
        "report_number": report_number, "services": services, "climate": climate,
        "progress": progress, "field_records": field_records, "photos": photos,
        "nonconformities": nonconformities, "environment": environment, "sst": sst,
        "meetings": meetings, "risks": risks, "interferences": interferences,
        "annexes": annexes, "closing_term": closing_term,
    }

    manifest["discovery"]["warnings"] = warnings
    return manifest, report_data, prov
