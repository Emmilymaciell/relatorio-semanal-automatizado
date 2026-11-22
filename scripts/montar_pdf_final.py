# -*- coding: utf-8 -*-
r"""
scripts/montar_pdf_final.py - Fase 3.3B (Checkpoint 3.3B): montagem final,
paginacao/Termo e validacao do PDF completo.

Le o que a Fase 3.3A ja deixou pronto em output/shadow/<obra>_<inicio>_<fim>/
(shadow_final_candidate_v{versao}.docx/.pdf, report_data.json, manifest.yaml,
assembly_plan.json) e faz o que so pode ser feito DEPOIS que o corpo existe
em PDF:

  1. Gera o placeholder A0 do Diagrama Unifilar (fonte real e' um .xlsx,
     Excel COM nao funciona aqui -- LIMITATION documentada, nunca inventa
     conteudo de diagrama).
  2. Monta o PDF completo: corpo -> unifilar (placeholder) -> diario, na
     ordem do assembly_plan.json (nunca por posicao hardcoded).
  3. Conta as paginas reais do PDF montado (nunca usa a contagem do golden).
  4. Roda a mecanica de paginacao do Termo (scripts/gerador/termo_paginacao.py)
     em 2 passagens reais: PASS 1 conta o PDF montado com o Termo ainda
     placeholder; PASS 2 aplica um patch DIRETO no PDF (fitz/redaction, sem
     Word/Excel/COM) com o numero real e reconta -- convergencia genuina,
     nao assumida.
  5. Atualiza o Termo dentro do .docx (python-docx), pronto para uma futura
     reconversao manual pelo Word.
  6. Roda o validador (scripts/validar_relatorio_semanal.py) sobre o PDF
     FINAL MONTADO (nao so o corpo).
  7. Avalia os 12 critical gates e grava
     output/shadow/<...>/critical_gates_v{versao}.json.

Uso:
    python scripts/montar_pdf_final.py --obra RD103 --inicio 2026-08-29 --fim 2026-09-04 --versao v0.1
"""
import os
import sys
import re
import json
import shutil
import argparse
import datetime as dt

import yaml

AQUI = os.path.dirname(os.path.abspath(__file__))
if AQUI not in sys.path:
    sys.path.insert(0, AQUI)
import prog_util as U  # noqa: E402

from gerador import config as C  # noqa: E402
from gerador import content_builder as CB  # noqa: E402
from gerador import termo_paginacao as TP  # noqa: E402
from gerador import pdf_final_assembly as PA  # noqa: E402
from gerador import termo_docx_patch as TDP  # noqa: E402
from gerador import compare as CMP  # noqa: E402

import validar_relatorio_semanal as VR  # noqa: E402


def _outdir(obra, di, df):
    return os.path.join(U.OUT, "shadow", "%s_%s_%s" % (obra, di.isoformat(), df.isoformat()))


# -----------------------------------------------------------------------------
# 1+2+3+4. montagem + paginacao do Termo (2 passagens reais)
# -----------------------------------------------------------------------------
def montar_e_paginar(outdir, versao, report_data, assembly_plan, obra):
    docx_path = os.path.join(outdir, "shadow_final_candidate_%s.docx" % versao)
    body_pdf_original = os.path.join(outdir, "shadow_final_candidate_%s.pdf" % versao)
    if not os.path.exists(body_pdf_original):
        raise RuntimeError("corpo em PDF não encontrado: %s -- rode a conversão manual "
                           "Word (DOCX->PDF) antes de montar o final" % body_pdf_original)

    external = assembly_plan["external_pdf_annexes"]
    unifilar_cfg = next(a for a in external if a["anexo_id"] == "diagrama_unifilar")
    diario_cfg = next(a for a in external if a["anexo_id"] == "diario_de_obras")

    diario_path = diario_cfg["source_path"]
    if not diario_path or not U.exists(diario_path):
        raise RuntimeError("Diário de Obras não encontrado na fonte: %s" % diario_path)

    anexos_dir = os.path.join(outdir, "anexos_externos")
    os.makedirs(anexos_dir, exist_ok=True)

    # Fase 3.4/v0.2: tenta a prancha REAL (UnifilarRenderer, le a propria
    # planilha via openpyxl+Pillow) e so cai no placeholder textual A0/A3
    # (Fase 3.3B) se a renderizacao real falhar por qualquer motivo -- nunca
    # quebra a montagem. `papel_pdf` vem do proprio assembly_plan.json
    # (road_registry.yaml#formato_unifilar), nunca de um `if obra ==` aqui.
    unifilar_out = os.path.join(anexos_dir, "diagrama_unifilar.pdf")
    figuras_dir = os.path.join(outdir, "figuras")
    resultado_unifilar = PA.gerar_anexo_unifilar(
        unifilar_out,
        caminho_xlsx=unifilar_cfg["source_path"],
        papel_pdf=unifilar_cfg.get("papel_pdf") or "A0 paisagem",
        fonte_esperada=unifilar_cfg["source_path"],
        figuras_dir=figuras_dir,
        motivo_fallback=(
            "UNIFILAR %s.xlsx não pôde ser renderizado a partir da própria planilha "
            "neste ambiente (openpyxl/Pillow) e nenhum PDF da semana 49 foi localizado "
            "na árvore da obra nem no compartilhado (só PDFs de relatórios mensais/"
            "semanas antigas)."
            % (obra[2:] if obra.startswith("GO") else obra)))

    body_pass1 = os.path.join(outdir, "body_pass1.pdf")
    shutil.copyfile(body_pdf_original, body_pass1)
    n_corpo = PA.contar_paginas(body_pass1)
    n_unifilar = PA.contar_paginas(unifilar_out)
    n_diario = PA.contar_paginas(diario_path)

    ordem = [body_pass1, unifilar_out, diario_path]

    info = C.obra_info(obra)
    road = report_data["road"]
    contract = report_data["contract"]
    period = report_data["period"]
    cidade = C.road_registry().get("meta", {}).get("cidade_emissao", "___")

    estado = {}

    def _texto_termo(n_paginas):
        return CB._fmt(
            "s10_termo",
            obra_sem_hifen=road.get("rodovia_sem_hifen") or info.get("rodovia_sem_hifen", "___"),
            numero="%d" % report_data["report_number"],
            data_inicio=CB._data_br(period["start"]), data_fim=CB._data_br(period["end"]),
            contrato_obra=contract.get("contrato_obra") or "___",
            executora_contratual=road.get("empresa_executora_contrato") or "___",
            n_paginas=str(n_paginas),
            n_paginas_extenso=CB.por_extenso(n_paginas, feminino=True),
            cidade=cidade,
            data_emissao_extenso=CB._data_por_extenso(dt.date.today()))

    def contador():
        n = estado.get("n", 0) + 1
        estado["n"] = n
        if n == 1:
            p = os.path.join(outdir, "merged_pass1.pdf")
            PA.montar_pdf(ordem, p)
            estado["merged_pass1"] = p
            c = PA.contar_paginas(p)
            estado["n_pass1"] = c
            return c
        elif n == 2:
            n_prev = estado["n_pass1"]
            texto_completo = _texto_termo(n_prev)
            estado["texto_final_completo"] = texto_completo
            partes = texto_completo.split("\n\n")
            texto_par = partes[0].strip()
            texto_fech = partes[1].strip() if len(partes) > 1 else ""
            p2 = os.path.join(outdir, "merged_pass2.pdf")
            resultado_patch = PA.patch_termo_pdf(
                estado["merged_pass1"], p2, texto_paragrafo=texto_par, texto_fechamento=texto_fech)
            estado["patch"] = resultado_patch
            if not resultado_patch["ok"]:
                return None
            estado["merged_pass2"] = p2
            return PA.contar_paginas(p2)
        else:
            caminho = estado.get("merged_pass2") or estado.get("merged_pass1")
            return PA.contar_paginas(caminho)

    resultado_tp = TP.rodar(contador)

    return {
        "docx_path": docx_path,
        "body_pdf_original": body_pdf_original,
        "unifilar_out": unifilar_out,
        "resultado_unifilar": resultado_unifilar,
        "diario_path": diario_path,
        "n_corpo": n_corpo,
        "n_unifilar": n_unifilar,
        "n_diario": n_diario,
        "resultado_tp": resultado_tp,
        "estado": estado,
    }


# -----------------------------------------------------------------------------
# 5. critical gates
# -----------------------------------------------------------------------------
def _normalizar_anexo(s):
    return CMP.normalizar_anexo_titulo(s)


_CATEGORIAS_ANEXO = (
    (r"^pa\b|^pa-|parecer", "parecer_tecnico"),
    (r"\bata\b|reuniao", "ata_reuniao"),
    (r"ambiental", "acompanhamento_ambiental"),
    (r"saude|seguranca|\bsst\b", "acompanhamento_sst"),
    (r"^nc\b|^nc-|nao conformidade", "nao_conformidade"),
    (r"unifilar", "diagrama_unifilar"),
    (r"diario", "diario_de_obras"),
    (r"circular", "circular_tecnica"),
    (r"desvio", "desvio_cronograma"),
)


def _categoria_anexo(titulo_ou_id):
    n = _normalizar_anexo(titulo_ou_id)
    for padrao, cat in _CATEGORIAS_ANEXO:
        if re.search(padrao, n):
            return cat
    return n


def _colapsar(seq):
    """Funde repeticoes consecutivas: o golden pode ter um anexo por NC (NC-022,
    NC-023) e o shadow um anexo 'Não Conformidade' agrupado."""
    out = []
    for x in seq:
        if not out or out[-1] != x:
            out.append(x)
    return out


def golden_titulos_fotografico(full_text):
    """Subsecoes 5.2.x do golden. O texto vem do DOCX SEM a numeracao automatica:
    cada subsecao e' a linha imediatamente antes de 'Figura N - Relatório fotográfico'."""
    return [t.strip().rstrip(":").strip() for t in re.findall(
        r"^([^\n]{2,60})\nFigura\s+\d+\s*[-–]\s*Relat[óo]rio fotogr[áa]fico",
        full_text, re.MULTILINE | re.IGNORECASE)]


def golden_n_nc_prog(full_text):
    """Quantas NC o golden declara emitidas pelo Programa no periodo (secao 8.1).
    0 se a 8.1 diz 'não houve'; senao, NC-NNN distintos citados; None se a 8.1 nao
    for localizada. A ULTIMA ocorrencia e' o corpo (o sumario pode repetir os titulos)."""
    ocorr = list(re.finditer(
        r"N[ãa]o conformidades emitidas pel[ao] Programa no per[íi]odo:?(.*?)"
        r"N[ãa]o conformidades emitidas pel[ao] Órgão Contratante",
        full_text, re.IGNORECASE | re.DOTALL))
    if not ocorr:
        return None
    corpo = ocorr[-1].group(1)
    if re.search(r"n[ãa]o houve", corpo, re.IGNORECASE):
        return 0
    return len(set(re.findall(r"NC-\d+", corpo)))


def _aplicar_excecoes(gates, obra, inicio):
    """config/gate_exceptions.yaml: divergencia golden x shadow com causa CONFIRMADA
    fora do motor (GOLDEN_DEFECT, SOURCE_LIMITATION). O gate continua reprovado na
    comparacao bruta; o status vira CRITICAL_GATE_PASS_COM_EXCECAO e a evidencia
    da excecao entra no registro. Sem entrada no arquivo, nada muda."""
    p = os.path.join(U.PROJ, "config", "gate_exceptions.yaml")
    if not os.path.exists(p):
        return gates
    with open(p, encoding="utf-8") as fh:
        excecoes = (yaml.safe_load(fh) or {}).get("excecoes") or []
    for g in gates:
        if g["status"] != "CRITICAL_GATE_FAIL":
            continue
        for e in excecoes:
            if (e.get("obra") == obra and str(e.get("inicio")) == inicio
                    and g["id"] in (e.get("gates") or [])):
                g["status"] = "CRITICAL_GATE_PASS_COM_EXCECAO"
                g["excecao"] = {k: e.get(k) for k in ("id", "classe", "nivel", "evidencia")}
                break
    return gates


def avaliar_critical_gates(report_data, manifest, montagem, n_final, final_pdf_path):
    golden = CMP.carregar_golden(manifest)
    golden_fatos = CMP.extrair_fatos_golden(golden) if golden and not golden.get("erro") else {}

    road = report_data["road"]
    contract = report_data["contract"]
    period = report_data["period"]

    gates = []

    def add(id_, nome, ok, evidencia, na=False):
        status = "N/A" if na else ("CRITICAL_GATE_PASS" if ok else "CRITICAL_GATE_FAIL")
        gates.append({"id": id_, "gate": nome, "status": status, "evidencia": evidencia})

    # 1. rodovia
    add("G01", "rodovia", golden_fatos.get("rodovia") == road.get("rodovia"),
        "shadow=%r golden=%r" % (road.get("rodovia"), golden_fatos.get("rodovia")))

    # 2. periodo
    periodo_shadow = "%s a %s" % (CB._data_br(period["start"]), CB._data_br(period["end"]))
    add("G02", "período", golden_fatos.get("periodo") == periodo_shadow,
        "shadow=%r golden=%r" % (periodo_shadow, golden_fatos.get("periodo")))

    # 3. numero
    numero_shadow = "%03d" % report_data["report_number"]
    add("G03", "número do relatório", golden_fatos.get("relatorio") == numero_shadow,
        "shadow=%r golden=%r" % (numero_shadow, golden_fatos.get("relatorio")))

    # 4. contrato
    add("G04", "contrato", golden_fatos.get("contrato_obra") == contract.get("contrato_obra"),
        "shadow=%r golden=%r" % (contract.get("contrato_obra"), golden_fatos.get("contrato_obra")))

    # 5. trecho
    trecho_golden_norm = U.norm(golden_fatos.get("trecho") or "")
    trecho_shadow_norm = U.norm(road.get("trecho") or "")
    add("G05", "trecho", bool(trecho_golden_norm) and trecho_golden_norm == trecho_shadow_norm,
        "shadow=%r golden=%r" % (road.get("trecho"), golden_fatos.get("trecho")))

    # 6. avanco fisico -- golden e' grafico (imagem), nao texto extraivel do
    # .docx; a verificacao document a comparacao FEITA no Checkpoint 3.3A
    # (fonte: mesma planilha 'Gráfico - Avanço Global.xlsx' usada pelo
    # processo manual, contada pagina a pagina no golden) -- nao pode ser
    # re-extraida por regex aqui porque o golden nao traz esses numeros como
    # texto. Reportado como PASS com o metodo registrado, nunca inventado.
    # multiobra-r1: ate a v0.3 este gate comparava com 32,23/42,19 FIXOS (os valores
    # da RD-103 semana 49) -- passava na RD-103 por construcao e reprovava qualquer
    # outra obra. Agora: o valor tem de ter sido LIDO da fonte (nao nulo), ser
    # coerente (0 < concluido <= concluido+atacado <= 1) e, quando o golden trouxer
    # os percentuais como texto, bater com eles. No v3 o golden os traz so no
    # grafico (imagem): sem texto, a comparacao com o golden nao e' possivel e o gate
    # se apoia na fonte, com o metodo registrado.
    prog = report_data["progress"]
    concl = prog.get("global_concluido")
    concl_atacado = prog.get("global_concluido_atacado")
    coerente = (concl is not None and concl_atacado is not None
                and 0 < concl <= concl_atacado <= 1)
    gtxt = (golden.get("full_text") if golden else "") or ""
    pct_golden = []
    if coerente:
        for v in (concl, concl_atacado):
            alvo = ("%.2f" % (v * 100)).replace(".", ",")
            pct_golden.append(alvo in gtxt)
    m_av = re.search(r"avan[çc]o global[^%]{0,160}?(\d{1,3},\d{2})\s*%", gtxt, re.IGNORECASE)
    if m_av and coerente:
        avanco_ok = all(pct_golden)
        metodo = "golden traz percentual em texto (%s%%): comparado" % m_av.group(1)
    else:
        avanco_ok = coerente
        metodo = ("golden traz o avanço só no gráfico (imagem): gate verifica leitura da "
                  "fonte e coerência concluído <= concluído+atacado")
    add("G06", "avanço físico (concluído / concluído+atacado)", avanco_ok,
        "shadow=%s / %s -- %s; fonte=%s"
        % ("%.2f%%" % (concl * 100) if concl is not None else None,
           "%.2f%%" % (concl_atacado * 100) if concl_atacado is not None else None,
           metodo, prog.get("_status") or "lido"))

    # 7. servicos -- compara pelo NOME CANONICO ou pela SIGLA conhecida
    # (config/services_taxonomy.yaml#codigos_ficha_de_campo), nunca so pela
    # 1a palavra: "Obras de Arte Especiais" (shadow) == "OAE" (golden,
    # abreviado) e' a MESMA equivalencia ja documentada em
    # comparison_rules.yaml (regra "servicos-sigla-equivalencia"), nao uma
    # divergencia de conteudo.
    servicos_shadow = [s["canonical_name"] for s in report_data["services"]]
    servicos_golden_frase = golden_fatos.get("servicos_s1_frase") or ""
    frase_norm = U.norm(servicos_golden_frase)
    taxo = C.services_taxonomy()
    sigla_por_fullname_norm = {U.norm(v): k for k, v in
                               (taxo.get("codigos_ficha_de_campo") or {}).items()}

    def _servico_no_golden(nome_canonico):
        n = U.norm(nome_canonico)
        if n and n in frase_norm:
            return True
        sigla = sigla_por_fullname_norm.get(n)
        return bool(sigla and re.search(r"\b%s\b" % re.escape(sigla.lower()), frase_norm))

    # multiobra-r1: o golden declara os servicos em DOIS lugares -- a frase da
    # Apresentacao e as subsecoes 5.2.x do fotografico. Na RD-101/48 a frase omite
    # "Obras Complementares" mas a 5.2.6 existe: o servico consta do golden. Vale a
    # uniao; e, no sentido inverso, toda subsecao 5.2.x do golden tem de existir no
    # shadow (antes so se checava shadow -> golden).
    # o texto do golden vem do DOCX, SEM a numeracao automatica: cada subsecao do
    # fotografico e' a linha imediatamente antes de "Figura N - Relatório fotográfico"
    titulos_52 = golden_titulos_fotografico((golden.get("full_text") if golden else "") or "")
    frase_norm = frase_norm + " " + " ".join(U.norm(t) for t in titulos_52)
    nomes_shadow_norm = set()
    for s in servicos_shadow:
        nomes_shadow_norm.add(U.norm(s))
        if sigla_por_fullname_norm.get(U.norm(s)):
            nomes_shadow_norm.add(sigla_por_fullname_norm[U.norm(s)].lower())
    faltando = [s for s in servicos_shadow if not _servico_no_golden(s)]
    sobrando_golden = [t for t in titulos_52 if U.norm(t) not in nomes_shadow_norm]
    add("G07", "serviços executados (frase + subseções 5.2.x)", not faltando and not sobrando_golden,
        "shadow=%s golden_frase=%r golden_5.2.x=%s%s%s" % (
            servicos_shadow, servicos_golden_frase, titulos_52,
            (" -- shadow sem par no golden=%s" % faltando) if faltando else "",
            (" -- golden sem par no shadow=%s" % sobrando_golden) if sobrando_golden else ""))

    # 8. NC emitidas no periodo
    # multiobra-r1: ate a v0.3 o gate exigia ZERO NC (o caso da RD-103/49) e o regex
    # "nao houve emissao" casava tambem a frase da Órgão Contratante (8.2). Agora compara a
    # CONTAGEM do shadow com a da secao 8.1 (Programa) do golden.
    nc_periodo = report_data["nonconformities"].get("issued_in_period") or []
    n_golden = golden_n_nc_prog((golden.get("full_text") if golden else "") or "")
    add("G08", "NC emitidas no período (Programa, seção 8.1)",
        n_golden is not None and len(nc_periodo) == n_golden,
        "shadow=%d %s golden_8.1=%s" % (
            len(nc_periodo), [n.get("documento") for n in nc_periodo], n_golden))

    # 9. PA da semana
    pa_shadow = next((a.get("pa_numero") for a in report_data["annexes"]
                      if a["id"] == "parecer_tecnico"), None)
    pa_golden_txt = golden.get("full_text", "") if golden else ""
    add("G09", "PA da semana", bool(pa_shadow) and pa_shadow in pa_golden_txt,
        "shadow=%r encontrado_no_texto_golden=%s" % (pa_shadow, pa_shadow in pa_golden_txt if pa_shadow else False))

    # 10. anexos -- multiobra-r1: ate a v0.3 exigia exatamente 5 (o caso da
    # RD-103/49). Agora: o multiconjunto de CATEGORIAS de anexo do shadow tem de ser
    # igual ao do golden (numero do PA e' o G09; numero de NC/ata nao entra aqui).
    anexos_shadow_ids = [a["id"] for a in report_data["annexes"] if a.get("presente")]
    cat_golden = _colapsar([_categoria_anexo(t) for _n, t in golden_fatos.get("anexos", [])])
    cat_shadow = _colapsar([_categoria_anexo(a["id"]) for a in report_data["annexes"]
                            if a.get("presente")])
    add("G10", "anexos (mesmas categorias do golden)",
        bool(cat_golden) and sorted(cat_golden) == sorted(cat_shadow),
        "shadow=%s golden=%s" % (cat_shadow, cat_golden))

    # 11. ordem dos anexos -- compara a ordem shadow (report_data['annexes'],
    # a mesma usada por content_builder/assembly_model) com a ordem golden
    # (extraida da lista real "Anexo N - Titulo" do golden)
    # titulo exibido de cada anexo -- MESMA composicao de
    # content_builder._secao11_anexos (titulo + " PA-NNN" so para o parecer),
    # nunca o `id` interno cru (que nao carrega "PA-254" e faria "parecer
    # tecnico" != "pa-254" por um bug de comparacao, nao por diferenca real).
    def _titulo_anexo_shadow(a):
        titulo = CB._TITULOS_ANEXO.get(a["id"], a["id"])
        if a["id"] == "parecer_tecnico" and a.get("pa_numero"):
            titulo = "%s %s" % (titulo, a["pa_numero"])
        return titulo

    # multiobra-r1: ordem por CATEGORIA ("nc" == "nc-022"; o golden numera a NC no
    # titulo e o shadow nao) -- antes comparava o titulo normalizado inteiro.
    anexos_golden_ordem = _colapsar([_categoria_anexo(t) for _n, t in golden_fatos.get("anexos", [])])
    anexos_shadow_ordem = _colapsar([_categoria_anexo(_titulo_anexo_shadow(a))
                                     for a in report_data["annexes"]])
    add("G11", "ordem dos anexos", bool(anexos_golden_ordem)
        and anexos_golden_ordem == anexos_shadow_ordem,
        "shadow=%s golden=%s" % (anexos_shadow_ordem, anexos_golden_ordem))

    # 12. Termo / paginacao
    tp = montagem["resultado_tp"]
    termo_ok = (tp.resultado == "ESTAVEL" and n_final is not None)
    add("G12", "Termo de Encerramento / paginação real", termo_ok,
        "resultado=%s n_paginas=%s passagens=%s" % (tp.resultado, tp.n_paginas, tp.passagens))

    return _aplicar_excecoes(gates, (manifest.get("report") or {}).get("road_key"),
                             str(period.get("start")))


# -----------------------------------------------------------------------------
# 6. validador completo sobre o PDF final
# -----------------------------------------------------------------------------
def _resolver_versao_relatorio(cfg, data_inicio):
    d = dt.date.fromisoformat(data_inicio)
    for vid in ("v3", "v2", "v1"):
        v = cfg["report_versions"][vid]
        ini = v["effective_from"]
        fim = v.get("effective_until")
        if d >= ini and (fim is None or d <= fim):
            return v
    return cfg["report_versions"][cfg["resolucao"]["padrao_se_indeterminado"]]


def validar_pdf_final(pdf_path, obra, report_data, di, df):
    cfg = yaml.safe_load(open(os.path.join(U.PROJ, "config", "report_versions.yaml"),
                              encoding="utf-8"))
    reg = yaml.safe_load(open(os.path.join(U.PROJ, "config", "road_registry.yaml"),
                              encoding="utf-8"))
    pdf = VR.ler_pdf(pdf_path)
    ver = _resolver_versao_relatorio(cfg, di.isoformat())
    ctx = {"obra": obra, "numero": report_data["report_number"],
          "data_inicio": di.isoformat(), "data_fim": df.isoformat(), "pdf": pdf, "versao": ver,
          "ancora": str(reg["obras"][obra].get("ancora_semana_1"))}
    resultados = VR.validar(ctx, cfg)
    return {"n_paginas": pdf["n_paginas"], "formatos": dict(pdf["formatos"]),
           "anexos_detectados": pdf["anexos"], "resultados": resultados}


# -----------------------------------------------------------------------------
def main():
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    ap = argparse.ArgumentParser()
    ap.add_argument("--obra", required=True, choices=U.OBRAS)
    ap.add_argument("--inicio", required=True)
    ap.add_argument("--fim", required=True)
    ap.add_argument("--versao", default="v0.1")
    args = ap.parse_args()

    di = dt.date.fromisoformat(args.inicio)
    df = dt.date.fromisoformat(args.fim)
    outdir = _outdir(args.obra, di, df)

    with open(os.path.join(outdir, "manifest.yaml"), encoding="utf-8") as fh:
        manifest = yaml.safe_load(fh)
    with open(os.path.join(outdir, "report_data.json"), encoding="utf-8") as fh:
        report_data = json.load(fh)
    with open(os.path.join(outdir, "assembly_plan.json"), encoding="utf-8") as fh:
        assembly_plan = json.load(fh)

    print("=" * 78)
    print("MONTAGEM FINAL  %s  %s..%s  (versão %s)" % (args.obra, di, df, args.versao))
    print("=" * 78)

    montagem = montar_e_paginar(outdir, args.versao, report_data, assembly_plan, args.obra)
    tp = montagem["resultado_tp"]
    estado = montagem["estado"]

    print("corpo (body_pass1.pdf): %d páginas" % montagem["n_corpo"])
    print("unifilar (%s): %d página -- %s"
         % (montagem["resultado_unifilar"]["metodo"], montagem["n_unifilar"],
            montagem["resultado_unifilar"].get("note") or montagem["resultado_unifilar"].get("motivo")))
    print("diário de obras (fonte real): %d páginas" % montagem["n_diario"])
    print("termo_paginacao.rodar(): %s" % tp.to_dict())

    n_final = tp.n_paginas if tp.resultado == "ESTAVEL" else estado.get("n_pass1")
    final_pdf = estado.get("merged_pass2") or estado.get("merged_pass1")
    final_pdf_nome = os.path.join(outdir, "shadow_final_candidate_%s_MONTADO.pdf" % args.versao)
    shutil.copyfile(final_pdf, final_pdf_nome)
    print("PDF final montado: %s (fonte: %s)" % (final_pdf_nome, os.path.basename(final_pdf)))

    # DOCX sempre atualizado com o numero real conhecido (n_pass1), mesmo que
    # o patch cosmetico do PDF (pass 2) nao tenha convergido.
    n_para_docx = estado.get("n_pass1")
    texto_docx = estado.get("texto_final_completo")
    if texto_docx is None:
        # pass 2 nao rodou (nao deveria acontecer, TP.rodar sempre chama pelo
        # menos 2x quando ha PDF) -- recalcula aqui por seguranca, nunca
        # deixa o DOCX sem atualizar por falta desse texto.
        info = C.obra_info(args.obra)
        road = report_data["road"]; contract = report_data["contract"]; period = report_data["period"]
        cidade = C.road_registry().get("meta", {}).get("cidade_emissao", "___")
        texto_docx = CB._fmt(
            "s10_termo", obra_sem_hifen=road.get("rodovia_sem_hifen") or info.get("rodovia_sem_hifen", "___"),
            numero="%d" % report_data["report_number"], data_inicio=CB._data_br(period["start"]),
            data_fim=CB._data_br(period["end"]), contrato_obra=contract.get("contrato_obra") or "___",
            executora_contratual=road.get("empresa_executora_contrato") or "___",
            n_paginas=str(n_para_docx), n_paginas_extenso=CB.por_extenso(n_para_docx, feminino=True),
            cidade=cidade, data_emissao_extenso=CB._data_por_extenso(dt.date.today()))

    docx_result = TDP.atualizar_termo_docx(montagem["docx_path"], texto_final=texto_docx)
    print("DOCX atualizado: %s" % docx_result)

    # validador completo sobre o PDF FINAL MONTADO
    validacao = validar_pdf_final(final_pdf_nome, args.obra, report_data, di, df)
    p_val = os.path.join(outdir, "validacao_pdf_final_%s.json" % args.versao)
    U.assert_read_only(p_val)
    with open(p_val, "w", encoding="utf-8") as fh:
        json.dump(validacao, fh, ensure_ascii=False, indent=2, default=str)
    print("validação: %s" % p_val)
    for r in validacao["resultados"]:
        print("  %-4s %-55s %-8s %s" % (r["id"], r["regra"][:55], r["status"], r["detalhe"]))

    # critical gates
    gates = avaliar_critical_gates(report_data, manifest, montagem, n_final, final_pdf_nome)
    p_gates = os.path.join(outdir, "critical_gates_%s.json" % args.versao)
    U.assert_read_only(p_gates)
    with open(p_gates, "w", encoding="utf-8") as fh:
        json.dump({"gerado_em": dt.datetime.now().astimezone().isoformat(timespec="seconds"),
                  "obra": args.obra, "periodo": {"start": str(di), "end": str(df)},
                  "n_paginas_final": n_final, "gates": gates,
                  # v0.3: o unifilar pode cair no placeholder A0 SILENCIOSAMENTE
                  # (mesmo formato de pagina, mesma contagem -- nada no PDF
                  # denuncia). O metodo usado fica registrado aqui para o
                  # regression guard poder exigir a prancha real.
                  "unifilar": {
                      "metodo": montagem["resultado_unifilar"].get("metodo"),
                      "ok": montagem["resultado_unifilar"].get("ok"),
                      "n_paginas": montagem["n_unifilar"],
                      "nota": (montagem["resultado_unifilar"].get("note")
                               or montagem["resultado_unifilar"].get("motivo")),
                  }}, fh, ensure_ascii=False, indent=2)
    print("critical gates: %s" % p_gates)
    for g in gates:
        print("  %-4s %-45s %s" % (g["id"], g["gate"], g["status"]))

    print("=" * 78)
    print("RESUMO")
    print("=" * 78)
    print("n_paginas_final (real, programático): %s" % n_final)
    print("termo: resultado=%s passagens=%s" % (tp.resultado, tp.passagens))
    print("PDF final: %s" % final_pdf_nome)
    n_fail_gates = sum(1 for g in gates if g["status"] == "CRITICAL_GATE_FAIL")
    print("critical gates FAIL: %d/12" % n_fail_gates)
    n_fail_val = sum(1 for r in validacao["resultados"] if r["status"] == "FAIL")
    print("validador FAIL: %d" % n_fail_val)
    return 0


if __name__ == "__main__":
    sys.exit(main())
