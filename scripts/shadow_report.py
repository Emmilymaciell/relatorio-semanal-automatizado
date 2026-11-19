# -*- coding: utf-8 -*-
r"""
scripts/shadow_report.py - CLI do relatorio-sombra (Fase 3, Slices 1 e 2).

    python scripts/shadow_report.py sources --obra RD103 --inicio 2026-08-29 --fim 2026-09-04
    python scripts/shadow_report.py data    --obra RD103 --inicio 2026-08-29 --fim 2026-09-04
    python scripts/shadow_report.py data    --obra RD103 --semana 49

`sources` roda o Slice 1 (source_discovery) e grava `manifest.yaml`.
`data`    roda os Slices 1+2 (canonical) e grava `manifest.yaml` +
          `report_data.json` + `report_provenance.json`.

Saida sempre em output/shadow/{OBRA}_{inicio}_{fim}/. Nunca escreve fora de
output/ e tmp/ (ver prog_util.assert_read_only).
"""
import os
import sys
import json
import argparse
import datetime as dt

import yaml

AQUI = os.path.dirname(os.path.abspath(__file__))
if AQUI not in sys.path:
    sys.path.insert(0, AQUI)
import prog_util as U  # noqa: E402

from gerador import config as C  # noqa: E402
from gerador import periodo as P  # noqa: E402
from gerador import compare as CMP  # noqa: E402
from gerador.source_discovery import construir_manifest  # noqa: E402
from gerador.canonical import construir  # noqa: E402
from gerador.content_builder import construir_documento  # noqa: E402
from gerador.renderers import docx_renderer  # noqa: E402
from gerador.renderers import docx_renderer_v2  # noqa: E402
from gerador import assembly_model as AM  # noqa: E402
from gerador import structural_diff as SD  # noqa: E402
from gerador import annex_embed as AE  # noqa: E402
from gerador import termo_paginacao as TP  # noqa: E402
from gerador import candidate_manifest as CM  # noqa: E402
from gerador import figuras_bridge as FB  # noqa: E402


def _outdir(obra, di, df):
    return os.path.join(U.OUT, "shadow", "%s_%s_%s" % (obra, di.isoformat(), df.isoformat()))


def _gravar_yaml(obj, path):
    U.assert_read_only(path)
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as fh:
        yaml.safe_dump(obj, fh, allow_unicode=True, sort_keys=False,
                       default_flow_style=False)
    return path


def _gravar_json(obj, path):
    U.assert_read_only(path)
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(obj, fh, ensure_ascii=False, indent=2, default=str)
    return path


def _resolver_periodo_args(args):
    return P.resolver_periodo(args.obra, inicio=args.inicio, fim=args.fim, semana=args.semana)


def cmd_sources(args):
    C.validar()
    di, df = _resolver_periodo_args(args)
    manifest = construir_manifest(args.obra, di, df)
    outdir = _outdir(args.obra, di, df)
    caminho = _gravar_yaml(manifest, os.path.join(outdir, "manifest.yaml"))

    print("=" * 74)
    print("SOURCES  %s  %s..%s" % (args.obra, di, df))
    print("=" * 74)
    print("manifest: %s" % caminho)
    print("report_number: %s (schema_version=%s)"
         % (manifest["report"]["report_number"], manifest["report"]["schema_version"]))
    print("field_records: %d/7 ok, faltando: %s"
         % (sum(1 for f in manifest["sources"]["field_records"] if f["status"] == "OK"),
            manifest["sources"]["field_records_missing_days"]))
    print("warnings: %d" % len(manifest["discovery"]["warnings"]))
    for w in manifest["discovery"]["warnings"]:
        print("  ~", w)
    print("not_found: %d" % len(manifest["discovery"]["not_found"]))
    for nf in manifest["discovery"]["not_found"]:
        print("  X", nf["fonte"], nf["status"])
    return 0


def cmd_data(args):
    C.validar()
    di, df = _resolver_periodo_args(args)
    manifest, report_data, prov = construir(args.obra, di, df)
    outdir = _outdir(args.obra, di, df)

    p_manifest = _gravar_yaml(manifest, os.path.join(outdir, "manifest.yaml"))
    p_data = _gravar_json(report_data, os.path.join(outdir, "report_data.json"))
    p_prov = prov.salvar(os.path.join(outdir, "report_provenance.json"))

    print("=" * 74)
    print("DATA  %s  %s..%s" % (args.obra, di, df))
    print("=" * 74)
    print("manifest:   %s" % p_manifest)
    print("data:       %s" % p_data)
    print("provenance: %s (%d chaves)" % (p_prov, len(prov.chaves())))
    print("report_number: %s" % report_data["report_number"])
    print("services: %s" % [s["canonical_name"] for s in report_data["services"]])
    print("climate.by_day: %s"
         % [(d["date"], d["clima"]) for d in report_data["climate"]["by_day"]])
    print("climate.days_without_work: %s" % report_data["climate"]["days_without_work"])
    print("progress: concluido=%s atacado=%s"
         % (report_data["progress"].get("global_concluido"),
            report_data["progress"].get("global_concluido_atacado")))
    print("nonconformities.issued_in_period: %d"
         % len(report_data["nonconformities"].get("issued_in_period") or []))
    print("environment.applies: %s" % report_data["environment"]["applies"])
    print("annexes: %s" % [a["id"] for a in report_data["annexes"]])
    print("warnings: %d" % len(manifest["discovery"]["warnings"]))
    for w in manifest["discovery"]["warnings"]:
        print("  ~", w)
    return 0


def cmd_build(args):
    """Slice 3: report_data.json -> document_model -> shadow_report.docx."""
    C.validar()
    di, df = _resolver_periodo_args(args)
    manifest, report_data, prov = construir(args.obra, di, df)
    outdir = _outdir(args.obra, di, df)

    documento = construir_documento(report_data)

    p_manifest = _gravar_yaml(manifest, os.path.join(outdir, "manifest.yaml"))
    p_data = _gravar_json(report_data, os.path.join(outdir, "report_data.json"))
    p_prov = prov.salvar(os.path.join(outdir, "report_provenance.json"))
    p_model = _gravar_json(documento.to_dict(), os.path.join(outdir, "document_model.json"))
    p_docx = docx_renderer.render(documento, os.path.join(outdir, "shadow_report.docx"))

    print("=" * 74)
    print("BUILD  %s  %s..%s" % (args.obra, di, df))
    print("=" * 74)
    print("manifest:        %s" % p_manifest)
    print("data:            %s" % p_data)
    print("provenance:      %s" % p_prov)
    print("document_model:  %s" % p_model)
    print("docx:            %s" % p_docx)
    print("sections: %d" % len(documento.sections))
    print("status por bloco: %s" % documento.status_counts())
    for nome, status in documento.section_status().items():
        print("  %-55s %s" % (nome, status))
    return 0


def cmd_build2(args):
    """Fase 3.2B, P1: report_data.json -> document_model (o MESMO do
    'build') -> shadow_report_v2.docx (design system global, ver
    docx_renderer_v2.py e config/visual_style_v3.yaml). NUNCA sobrescreve o
    shadow_report.docx (v1) -- grava sempre num arquivo "..._v2.docx"
    separado. Tambem grava renderer_v2_structural_diff.json (Golden x v1 x
    v2), lendo o golden a partir de tmp/ (nunca do SharePoint -- so leitura
    de copia local ja existente)."""
    C.validar()
    di, df = _resolver_periodo_args(args)
    manifest, report_data, prov = construir(args.obra, di, df)
    outdir = _outdir(args.obra, di, df)

    documento = construir_documento(report_data)
    p_docx_v2 = docx_renderer_v2.render(documento, os.path.join(outdir, "shadow_report_v2.docx"))

    p_v1 = os.path.join(outdir, "shadow_report.docx")
    p_diff = SD.gerar_diff(
        golden_path=manifest.get("golden_reference", {}).get("docx", {}).get("path"),
        v1_path=p_v1 if os.path.exists(p_v1) else None,
        v2_path=p_docx_v2,
        output_path=os.path.join(outdir, "renderer_v2_structural_diff.json"),
    )

    print("=" * 74)
    print("BUILD2 (P1 design system)  %s  %s..%s" % (args.obra, di, df))
    print("=" * 74)
    print("shadow_report_v2.docx: %s" % p_docx_v2)
    print("shadow_report.docx (v1, intacto): %s" % (p_v1 if os.path.exists(p_v1) else "AUSENTE -- rode 'build' antes"))
    print("renderer_v2_structural_diff.json: %s" % p_diff)
    return 0


def cmd_assembly(args):
    """Fase 3.2B, P0.2: monta e serializa o AssemblyPlan (main_document +
    external_pdf_annexes). NAO executa o merge com os PDFs externos nesta
    rodada -- so descreve o plano em assembly_plan.json. NAO renderiza nem
    sobrescreve nenhum .docx (isso e' 'build')."""
    C.validar()
    di, df = _resolver_periodo_args(args)
    manifest, report_data, prov = construir(args.obra, di, df)
    outdir = _outdir(args.obra, di, df)

    documento = construir_documento(report_data)
    plano = AM.montar_plano(
        report_id="RO_%03d" % report_data["report_number"],
        obra=args.obra,
        periodo={"start": str(di), "end": str(df)},
        sections=documento.sections,
        annexes=report_data["annexes"],
        output_docx_path=os.path.join(outdir, "shadow_report_v2.docx"),
    )
    p_plan = AM.serializar(plano, os.path.join(outdir, "assembly_plan.json"))

    print("=" * 74)
    print("ASSEMBLY  %s  %s..%s" % (args.obra, di, df))
    print("=" * 74)
    print("assembly_plan: %s" % p_plan)
    print("main_document.sections: %d" % len(plano.main_document.sections))
    print("embedded_annexes (EMBEDDED_IMAGE): %s"
         % [a.anexo_id for a in plano.main_document.embedded_annexes])
    print("external_pdf_annexes (MERGED_PDF): %s"
         % [(a.anexo_id, a.posicao, a.papel_pdf) for a in plano.external_pdf_annexes])
    print("merge_executed: %s" % plano.merge_executed)
    return 0


def cmd_candidate(args):
    """Fase 3.3 (FASE3_3_FINAL_CANDIDATE_SPEC.md): pipeline completo --

        canonical -> content_builder -> document_model -> docx_renderer_v2
        -> anexos embutidos (fitz 200dpi, Sec.4) -> Termo (mecanica de
        paginacao, Sec.5) -> manifesto de versao (Sec.6)
        -> shadow_final_candidate_v{versao}.docx

    NUNCA sobrescreve shadow_report.docx (v1) nem shadow_report_v2.docx --
    grava sempre em '..._final_candidate_v{versao}.docx', um arquivo
    proprio desta fase (P0/Sec.0 da spec)."""
    C.validar()
    di, df = _resolver_periodo_args(args)
    manifest, report_data, prov = construir(args.obra, di, df)
    outdir = _outdir(args.obra, di, df)
    versao = args.versao
    # multiobra-r1: o candidate reconstroi os dados em memoria, mas a montagem e os
    # critical gates (montar_pdf_final.py) leem manifest/report_data DO DISCO. Sem
    # regravar aqui, o DOCX nascia dos dados novos e os gates julgavam os antigos
    # (RD-102A r1: NC-047-1 no PDF, G08 com shadow=0).
    _gravar_yaml(manifest, os.path.join(outdir, "manifest.yaml"))
    _gravar_json(report_data, os.path.join(outdir, "report_data.json"))
    prov.salvar(os.path.join(outdir, "report_provenance.json"))

    # Sec.1 (contrato FiguraResult): tenta as figuras REAIS da Frente A
    # (scripts/gerador/figuras/) por import tolerante -- se o pacote (ou uma
    # figura especifica dele) ainda nao existir, ou a chamada falhar,
    # figuras_bridge.tentar() devolve None e o content_builder mantem o
    # placeholder de sempre. O CAMINHO de arquivo (planilha da ficha
    # acumulada) e' resolvido AQUI, no pipeline -- nunca dentro de
    # content_builder.py (que continua sem conhecer arquivo/planilha).
    figuras_dir = os.path.join(outdir, "figuras")
    caminho_ficha_acumulada = (
        ((manifest.get("sources") or {}).get("accumulated_sheet") or {}).get("path"))
    figura_ficha = FB.tentar(
        "ficha_acumulada", "gerar", caminho_ficha_acumulada=caminho_ficha_acumulada,
        road_key=args.obra, out_dir=figuras_dir)
    figura_avanco = FB.tentar(
        "avanco_grafico", "gerar", report_data=report_data, out_dir=figuras_dir)
    paineis_fotograficos = FB.tentar(
        "painel_fotografico", "montar", report_data=report_data, manifest=manifest,
        out_dir=figuras_dir)

    # Correção 3 (Fase 3.3): liga as figuras ambientais (6/7/8) e de NC
    # (9, Gráfico 2, Gráfico 3) que a Frente B tinha deixado como
    # placeholders -- os PNGs já existiam em output/.../figuras/ (Frente A
    # tinha rodado antes de a Frente A terminar de fato o pipeline de
    # ligação). Caminho/empresa/rodovia vêm SEMPRE do manifest/report_data
    # (nunca hardcoded aqui), mesma resolução usada por tests/test_figuras.py.
    caminho_licenca = (((manifest.get("sources") or {}).get("environmental") or {})
                       .get("licenca_ambiental_planilha") or {}).get("path")
    caminho_nc_tracker = ((manifest.get("sources") or {}).get("nc_tracker") or {}).get("path")
    rodovia = report_data["road"].get("rodovia")

    # ambiental.py v0.2: filtro por Rodovia, nao por empresa executora do
    # contrato (SHADOW_ERROR corrigido -- ver docstring de
    # gerador/figuras/ambiental.py e docs/fontes/dashboard_ambiental_powerbi.md).
    figura_ambiental_indices = FB.tentar(
        "ambiental", "gerar_indices", caminho_licenca_ambiental=caminho_licenca,
        rodovia=rodovia, out_dir=figuras_dir)
    figura_ambiental_condicionantes = FB.tentar(
        "ambiental", "gerar_condicionantes", caminho_licenca_ambiental=caminho_licenca,
        rodovia=rodovia, out_dir=figuras_dir)
    figura_ambiental_qrcode = FB.tentar("ambiental", "gerar_qrcode", out_dir=figuras_dir)

    figura_nc_notificacoes = FB.tentar(
        "nc_graficos", "gerar_figura_notificacoes", report_data=report_data, out_dir=figuras_dir)
    figura_nc_status = FB.tentar(
        "nc_graficos", "gerar_grafico_status", report_data=report_data, out_dir=figuras_dir)
    figura_nc_natureza = FB.tentar(
        "nc_graficos", "gerar_grafico_natureza", caminho_nc_tracker=caminho_nc_tracker,
        rodovia=rodovia, out_dir=figuras_dir)

    figuras_resolvidas = {
        "ficha_acumulada": figura_ficha,
        "avanco_grafico": figura_avanco,
        "painel_fotografico": paineis_fotograficos if isinstance(paineis_fotograficos, dict) else {},
        "ambiental": {
            "indices": figura_ambiental_indices,
            "condicionantes": figura_ambiental_condicionantes,
            "qrcode": figura_ambiental_qrcode,
        },
        "nc": {
            "notificacoes": figura_nc_notificacoes,
            "grafico_status": figura_nc_status,
            "grafico_natureza": figura_nc_natureza,
        },
    }

    modo = getattr(args, "modo", None)
    documento = construir_documento(report_data, figuras=figuras_resolvidas, modo=modo)
    print("generation_mode: %s" % report_data["metadata"]["generation_mode"])

    plano = AM.montar_plano(
        report_id="RO_%03d" % report_data["report_number"], obra=args.obra,
        periodo={"start": str(di), "end": str(df)}, sections=documento.sections,
        annexes=report_data["annexes"],
        output_docx_path=os.path.join(outdir, "shadow_final_candidate_%s.docx" % versao),
    )

    # Sec.4: anexos EMBEDDED_IMAGE (PA-254/Ambiental/SST) -- fitz 200dpi,
    # pagina cheia, um bloco por pagina REAL do PDF fonte, na ordem do
    # assembly_plan. Anexados na PROPRIA Section 11 (ANEXOS), depois da
    # lista -- content_builder.py nunca abre arquivo, so o pipeline faz isso.
    vs = C.visual_style()
    m_corpo = vs["margins_mm"]["body_a4"]
    usable_w = vs["page"]["a4_mm"]["width"] - m_corpo["left"] - m_corpo["right"]
    usable_h = vs["page"]["a4_mm"]["height"] - m_corpo["top"] - m_corpo["bottom"]
    anexos_dir = os.path.join(outdir, "anexos_embutidos")
    blocos_embutidos, resumo_anexos = AE.montar_blocos_embutidos(
        plano.main_document.embedded_annexes, anexos_dir, usable_w, usable_h)

    secao11 = next((s for s in documento.sections if s.id == "s11_anexos"), None)
    if secao11 is None:
        raise RuntimeError("document_model sem a secao 's11_anexos' -- nao ha onde embutir os anexos")
    secao11.blocks.extend(blocos_embutidos)

    # Sec.5: mecanica de paginacao do Termo -- SEM_PDF nesta rodada (nao ha
    # conversao DOCX->PDF disponivel; Word/Excel COM nao funcionam aqui,
    # ver Sec.0). O Termo continua PLACEHOLDER (content_builder ja implementa
    # isso); aqui so registramos o resultado da mecanica, auditavel, no
    # manifesto.
    resultado_termo = TP.rodar(None)

    p_docx = docx_renderer_v2.render(
        documento, os.path.join(outdir, "shadow_final_candidate_%s.docx" % versao))
    p_plan = AM.serializar(plano, os.path.join(outdir, "assembly_plan.json"))

    # Sec.6: manifesto de versao -- hashes de fonte e config, reproduzivel.
    manifesto = CM.montar(
        obra=args.obra, road_label=U.RODOVIA_LABEL[args.obra],
        periodo=(str(di), str(df)), report_data=report_data, manifest=manifest,
        extra_source_paths=[a.get("path") for a in (report_data.get("annexes") or [])],
        renderer_version="v2", candidate_version=versao,
        validation_result={
            "document_status_counts": documento.status_counts(),
            "document_section_status": documento.section_status(),
            "embedded_annexes": resumo_anexos,
            "termo_paginacao": resultado_termo.to_dict(),
            # v0.3: as contagens de foto por painel (9/12/7/9 na RD-103)
            # so existiam no LOG de painel_fotografico.montar(). O regression
            # guard tinha de reabrir a planilha de 160 MB para conferi-las, e
            # sem a flag --fotos reportava o invariante como NAO_VERIFICAVEL.
            # Ficam persistidas aqui, ao lado do status: e' rastreabilidade,
            # nao mudanca na selecao -- o numero e' o mesmo que a figura usou.
            "painel_fotografico_contagem": (
                {k: FB.campo(v, "provenance", {}).get("n_fotos")
                 for k, v in paineis_fotograficos.items()}
                if isinstance(paineis_fotograficos, dict) else None),
            "painel_fotografico_grid": (
                {k: FB.campo(v, "provenance", {}).get("grid")
                 for k, v in paineis_fotograficos.items()}
                if isinstance(paineis_fotograficos, dict) else None),
            "figuras_frente_a": {
                "ficha_acumulada": FB.campo(figura_ficha, "status") if figura_ficha else "MODULO_INDISPONIVEL",
                "avanco_grafico": FB.campo(figura_avanco, "status") if figura_avanco else "MODULO_INDISPONIVEL",
                "painel_fotografico": (
                    {k: FB.campo(v, "status") for k, v in paineis_fotograficos.items()}
                    if isinstance(paineis_fotograficos, dict) else "MODULO_INDISPONIVEL"),
                "ambiental": {
                    "indices": FB.campo(figura_ambiental_indices, "status")
                              if figura_ambiental_indices else "MODULO_INDISPONIVEL",
                    "condicionantes": FB.campo(figura_ambiental_condicionantes, "status")
                                     if figura_ambiental_condicionantes else "MODULO_INDISPONIVEL",
                    "qrcode": FB.campo(figura_ambiental_qrcode, "status")
                             if figura_ambiental_qrcode else "MODULO_INDISPONIVEL",
                },
                "nc": {
                    "notificacoes": FB.campo(figura_nc_notificacoes, "status")
                                   if figura_nc_notificacoes else "MODULO_INDISPONIVEL",
                    "grafico_status": FB.campo(figura_nc_status, "status")
                                     if figura_nc_status else "MODULO_INDISPONIVEL",
                    "grafico_natureza": FB.campo(figura_nc_natureza, "status")
                                       if figura_nc_natureza else "MODULO_INDISPONIVEL",
                },
            },
        },
    )
    p_manifest = _gravar_json(manifesto, os.path.join(outdir, "candidate_manifest_%s.json" % versao))

    print("=" * 74)
    print("CANDIDATE %s  %s  %s..%s" % (versao, args.obra, di, df))
    print("=" * 74)
    print("shadow_final_candidate_%s.docx: %s" % (versao, p_docx))
    for nome_v1_v2 in ("shadow_report.docx", "shadow_report_v2.docx"):
        caminho_v = os.path.join(outdir, nome_v1_v2)
        print("%-24s %s" % (nome_v1_v2 + ":",
                            ("intacto, mtime=%s" % dt.datetime.fromtimestamp(os.path.getmtime(caminho_v)))
                            if os.path.exists(caminho_v) else "AUSENTE (rode 'build'/'build2' antes)"))
    print("assembly_plan.json: %s" % p_plan)
    print("candidate_manifest: %s" % p_manifest)
    print("figuras (Frente A, Sec.1): ficha_acumulada=%s avanco_grafico=%s painel_fotografico=%s"
         % (FB.campo(figura_ficha, "status") if figura_ficha else "MODULO_INDISPONIVEL",
            FB.campo(figura_avanco, "status") if figura_avanco else "MODULO_INDISPONIVEL",
            {k: FB.campo(v, "status") for k, v in paineis_fotograficos.items()}
            if isinstance(paineis_fotograficos, dict) else "MODULO_INDISPONIVEL"))
    print("figuras ambientais (Correção 3): indices=%s condicionantes=%s qrcode=%s"
         % (FB.campo(figura_ambiental_indices, "status") if figura_ambiental_indices else "MODULO_INDISPONIVEL",
            FB.campo(figura_ambiental_condicionantes, "status") if figura_ambiental_condicionantes else "MODULO_INDISPONIVEL",
            FB.campo(figura_ambiental_qrcode, "status") if figura_ambiental_qrcode else "MODULO_INDISPONIVEL"))
    print("figuras NC (Correção 3): notificacoes=%s grafico_status=%s grafico_natureza=%s"
         % (FB.campo(figura_nc_notificacoes, "status") if figura_nc_notificacoes else "MODULO_INDISPONIVEL",
            FB.campo(figura_nc_status, "status") if figura_nc_status else "MODULO_INDISPONIVEL",
            FB.campo(figura_nc_natureza, "status") if figura_nc_natureza else "MODULO_INDISPONIVEL"))
    print("anexos embutidos (Sec.4 / Correção 1): %s" % resumo_anexos)
    print("termo (Sec.5): %s" % resultado_termo.to_dict())
    print("document.status_counts(): %s" % documento.status_counts())
    return 0


def cmd_compare(args):
    """Slice 3: shadow x golden -> comparison.json + comparison_report.md."""
    C.validar()
    di, df = _resolver_periodo_args(args)
    outdir = _outdir(args.obra, di, df)

    caminho_manifest = os.path.join(outdir, "manifest.yaml")
    caminho_data = os.path.join(outdir, "report_data.json")
    if not (os.path.exists(caminho_manifest) and os.path.exists(caminho_data)):
        raise RuntimeError("rode 'build' (ou 'data') antes de 'compare' -- faltam manifest.yaml/"
                           "report_data.json em %s" % outdir)
    with open(caminho_manifest, encoding="utf-8") as fh:
        manifest = yaml.safe_load(fh)
    with open(caminho_data, encoding="utf-8") as fh:
        report_data = json.load(fh)

    documento = construir_documento(report_data)
    resultado = CMP.comparar(report_data, documento, manifest)
    p_json = _gravar_json(resultado, os.path.join(outdir, "comparison.json"))

    md = CMP.gerar_relatorio_md(resultado)
    caminho_md = os.path.join(outdir, "comparison_report.md")
    U.assert_read_only(caminho_md)
    with open(caminho_md, "w", encoding="utf-8") as fh:
        fh.write(md)

    print("=" * 74)
    print("COMPARE  %s  %s..%s" % (args.obra, di, df))
    print("=" * 74)
    if resultado.get("erro"):
        print("ERRO:", resultado["erro"])
        return 1
    print("comparison.json:        %s" % p_json)
    print("comparison_report.md:   %s" % caminho_md)
    for classe, n in resultado["contagem_por_classe"].items():
        print("  %-26s %d" % (classe, n))
    return 0


def main():
    ap = argparse.ArgumentParser(description="Relatorio-sombra Programa (Fase 3, Slices 1-3)")
    sub = ap.add_subparsers(dest="comando", required=True)

    for nome, fn in (("sources", cmd_sources), ("data", cmd_data), ("build", cmd_build),
                     ("build2", cmd_build2), ("assembly", cmd_assembly), ("compare", cmd_compare),
                     ("candidate", cmd_candidate)):
        sp = sub.add_parser(nome)
        sp.add_argument("--obra", required=True, choices=U.OBRAS)
        sp.add_argument("--inicio", default=None, help="AAAA-MM-DD (sabado)")
        sp.add_argument("--fim", default=None, help="AAAA-MM-DD (opcional; default inicio+6)")
        sp.add_argument("--semana", type=int, default=None, help="numero da semana (alternativa a --inicio)")
        if nome == "candidate":
            sp.add_argument("--versao", default="v0.1", help="versao do candidate (default v0.1)")
            # v0.3: este script E' o gerador do SHADOW -- ele reconstroi uma
            # semana que ja teve relatorio emitido. Por isso o default e'
            # HISTORICAL_SHADOW, e nao a inferencia por data: a inferencia
            # olha "faz quanto tempo que a semana acabou", que e' um proxy
            # ruim. O que importa e' se o relatorio ja foi emitido uma vez --
            # se foi, a fonte de estado atual nao descreve mais aquela data.
            sp.add_argument("--modo", default="HISTORICAL_SHADOW",
                            choices=["CURRENT", "HISTORICAL_SHADOW"],
                            help="modo de geracao (default HISTORICAL_SHADOW; "
                                 "ver config/generation_modes.yaml)")
        sp.set_defaults(func=fn)

    args = ap.parse_args()
    try:
        return args.func(args)
    except (ValueError, RuntimeError) as e:
        print("ERRO:", e, file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
