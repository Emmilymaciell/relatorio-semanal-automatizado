# -*- coding: utf-8 -*-
r"""
scripts/gerador/compare.py - Fase 3, Slice 3: shadow x golden -> comparison.json
+ comparison_report.md (spec Sec.6).

ÚNICA camada do gerador autorizada a LER o golden (RO_NNN.n - <rodovia>.docx),
e só para CLASSIFICAR diferenças -- nunca como fonte de dado para o
document_model/content_builder. `assert_nao_e_golden()` continua valendo para
todo o resto do gerador; este modulo NAO a chama (e' a excecao documentada em
FASE3_SLICE_3_SPEC.md §0.5).

5 eixos: metadata, section, text, data, structure.
9 classes: EXACT, SEMANTICALLY_EQUIVALENT, EXPECTED_DIFFERENCE, GOLDEN_DEFECT,
SHADOW_ERROR, MISSING, EXTRA, HITL_REQUIRED, SOURCE_LIMITATION.

Classificacao de cada item: 1) tenta a regra explicita de
config/comparison_rules.yaml (so aplica se a `condicao` bater nos dados
REAIS); 2) senao, igualdade normalizada -> EXACT; 3) senao,
ausencia/excesso estrutural -> MISSING/EXTRA; 4) fallback -> HITL_REQUIRED
com motivo -- NUNCA SHADOW_ERROR por omissao de regra.
"""
import os
import re
import sys
import json
import unicodedata
import datetime as dt

import docx

AQUI = os.path.dirname(os.path.abspath(__file__))
SCRIPTS = os.path.dirname(AQUI)
if SCRIPTS not in sys.path:
    sys.path.insert(0, SCRIPTS)
import prog_util as U  # noqa: E402

from . import config as C  # noqa: E402

CLASSES_VALIDAS = (
    "EXACT", "SEMANTICALLY_EQUIVALENT", "EXPECTED_DIFFERENCE", "GOLDEN_DEFECT",
    "SHADOW_ERROR", "MISSING", "EXTRA", "HITL_REQUIRED", "SOURCE_LIMITATION",
)
EIXOS_VALIDOS = ("metadata", "section", "text", "data", "structure")


# -----------------------------------------------------------------------------
# leitura do golden -- SO AQUI (spec Sec.0.5)
# -----------------------------------------------------------------------------
def _ler_golden_docx(path):
    """Copia para tmp/ (nunca abre o original) e devolve {"paragraphs": [...],
    "tables": [[[cell,...], ...], ...], "full_text": "..."}."""
    if not path or not U.exists(path):
        return None
    local = U.copiar_para_tmp(path, "golden_compare", "golden.docx")
    d = docx.Document(local)
    paragrafos = [p.text for p in d.paragraphs if p.text.strip()]
    tabelas = [[[c.text for c in row.cells] for row in t.rows] for t in d.tables]
    return {"paragraphs": paragrafos, "tables": tabelas, "full_text": "\n".join(paragrafos)}


def carregar_golden(manifest):
    """`manifest`: dict do manifest.yaml (Slice 1). Devolve o texto/tabelas do
    golden .docx, ou None com aviso se nao existir/nao puder ser lido."""
    ref = (manifest or {}).get("golden_reference", {}).get("docx", {})
    path = ref.get("path")
    try:
        return _ler_golden_docx(path)
    except Exception as e:
        return {"erro": "%s: %s" % (type(e).__name__, e), "paragraphs": [], "tables": [], "full_text": ""}


# -----------------------------------------------------------------------------
# normalizacao + avaliacao de regra
# -----------------------------------------------------------------------------
def _norm(s):
    if s is None:
        return None
    s = unicodedata.normalize("NFKD", str(s))
    s = "".join(c for c in s if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", s).strip().lower().rstrip(":").strip()


def _carregar_regras():
    doc = C._carregar("comparison_rules.yaml")
    return {r["id"]: r for r in doc.get("regras", [])}


def _avaliar_condicao(condicao, golden, shadow, shadow_status=None):
    ns = {"golden": golden, "shadow": shadow, "shadow_status": shadow_status}
    try:
        return bool(eval(condicao, {"__builtins__": {}}, ns))  # noqa: S307 -- config confiavel, sem builtins
    except Exception:
        return False


def classificar(item, regras):
    """item: {"id","eixo","campo","golden","shadow","shadow_status","regra_id"}.
    Devolve (classe, evidencia_ou_motivo, regra_aplicada_id)."""
    golden, shadow = item.get("golden"), item.get("shadow")
    regra_id = item.get("regra_id")
    if regra_id:
        regra = regras.get(regra_id)
        if regra is None:
            raise KeyError("compare.py referencia regra_id %r ausente de "
                           "config/comparison_rules.yaml" % regra_id)
        if _avaliar_condicao(regra["condicao"], golden, shadow, item.get("shadow_status")):
            return regra["classe"], regra.get("evidencia"), regra["id"]

    if golden is not None and _norm(golden) == _norm(shadow):
        return "EXACT", None, None

    vazio = (None, "", [], {})
    if golden not in vazio and shadow in vazio:
        return "MISSING", "presente no golden, ausente no Shadow", None
    if shadow not in vazio and golden in vazio:
        return "EXTRA", "presente no Shadow, ausente no golden", None

    return ("HITL_REQUIRED",
           "sem regra em comparison_rules.yaml que bata nos dados reais -- "
           "requer confirmação humana (nunca classificado como SHADOW_ERROR por omissão)",
           None)


# -----------------------------------------------------------------------------
# extracao de fatos do golden (regex sobre o texto pleno + tabelas)
# -----------------------------------------------------------------------------
_RE_CAPA = {
    "rodovia": re.compile(r"RODOVIA:\s*(.+)"),
    "trecho": re.compile(r"TRECHO:\s*(.+)"),
    "extensao": re.compile(r"EXTENS[ÃA]O:\s*(.+)"),
    "codigo_sre": re.compile(r"C[ÓO]DIGO SRE:\s*(.+)"),
    "lote": re.compile(r"LOTE:\s*(.+)"),
    "relatorio": re.compile(r"Relat[óo]rio:\s*(\d+)"),
    "executora": re.compile(r"EMPRESA EXECUTORA:\s*(.+)"),
    "supervisora": re.compile(r"EMPRESA SUPERVISORA:\s*(.+)"),
    "periodo": re.compile(r"PER[ÍI]ODO:\s*(\d{2}/\d{2}/\d{4})\s*a\s*(\d{2}/\d{2}/\d{4})"),
}
_RE_INSTRUMENTO = re.compile(r"INSTRUMENTO DE CONVOCA[ÇC][ÃA]O\s*N[º°]\s*(\S+)", re.IGNORECASE)
_RE_CONTRATO_ASSINADO = re.compile(
    r"contrato\s*\(([\d/]+)[^)]*\)\s*(?:foi\s*)?assinado\s*em\s*(\d{2}/\d{2}/\d{4})", re.IGNORECASE)
_RE_OS = re.compile(r"ordem de servi[çc]o\s*([\d/]+)\s*emitida\s*em\s*(\d{2}/\d{2}/\d{4})", re.IGNORECASE)
_RE_SERVICOS_S1 = re.compile(
    r"atividades (?:de|referentes a)\s+(.+?)\.", re.IGNORECASE)
# 1o match no full_text == 1a ocorrencia no documento (paragrafos em ordem de
# leitura) == a frase da secao 1, sempre antes da repeticao na secao 4
_RE_NC_ATRASO = re.compile(r"em atraso\s+[a-zà-ú\s]+?\((\d{1,3})\)", re.IGNORECASE)
_RE_NC_ENCERRADAS = re.compile(r"\((\d{1,3})\)\s*encerradas", re.IGNORECASE)
_RE_LICENCA_NUM = re.compile(r"Licen[çc]a Ambiental [ÚU]nica\s*n[º°]?\s*(\S+)", re.IGNORECASE)
_RE_TERMO = re.compile(
    r"constitu[ií]do de\s*(\d+)\s*\(([^)]+)\)\s*p[áa]ginas", re.IGNORECASE)


def extrair_fatos_golden(golden):
    """Devolve um dict de fatos pontuais, extraidos por regex do texto do
    golden. Cada campo ausente fica None (nunca inventado)."""
    texto = golden.get("full_text", "") if golden else ""
    fatos = {}
    for chave, rx in _RE_CAPA.items():
        m = rx.search(texto)
        if chave == "periodo":
            fatos["periodo"] = "%s a %s" % m.groups() if m else None
        else:
            fatos[chave] = m.group(1).strip() if m else None

    m = _RE_INSTRUMENTO.search(texto)
    fatos["instrumento_convocacao"] = m.group(1) if m else None
    m = _RE_CONTRATO_ASSINADO.search(texto)
    fatos["contrato_obra"] = m.group(1) if m else None
    fatos["contrato_assinado_em"] = m.group(2) if m else None
    m = _RE_OS.search(texto)
    fatos["ordem_servico"] = m.group(1) if m else None
    fatos["ordem_servico_em"] = m.group(2) if m else None

    m = _RE_SERVICOS_S1.search(texto)
    fatos["servicos_s1_frase"] = m.group(1).strip() if m else None

    m = _RE_NC_ATRASO.search(texto)
    fatos["nc_em_atraso"] = int(m.group(1)) if m else None
    m = _RE_NC_ENCERRADAS.search(texto)
    fatos["nc_encerradas"] = int(m.group(1)) if m else None

    m = _RE_LICENCA_NUM.search(texto)
    fatos["licenca_numero"] = m.group(1).rstrip(",.") if m else None

    m = _RE_TERMO.search(texto)
    fatos["termo_n_paginas"] = m.group(1) if m else None
    fatos["termo_n_paginas_extenso"] = m.group(2) if m else None

    # tabela climatica: tabela cujo 1a celula da 1a linha comeca com "DATA"
    fatos["climate_by_day"] = {}
    for tabela in golden.get("tables", []):
        if tabela and tabela[0] and _norm(tabela[0][0]).startswith("data"):
            datas, dias, climas = tabela[0][1:], tabela[1][1:], tabela[2][1:]
            for d, c in zip(datas, climas):
                fatos["climate_by_day"][d] = c.strip()
            break

    # titulos de secao especificos (presenca literal no texto)
    for chave, alvo in (
        ("titulo_5_1", re.compile(
            r"^(?:Dados clim[áa]ticos|Condi[çc][õo]es clim[áa]ticas) do per[íi]odo:?\s*$", re.M)),
        ("titulo_9", re.compile(r"^RELATO DE (RISCOS E )?INTERFER[ÊE]NCIAS? EM CAMPO\s*$", re.M)),
    ):
        m = alvo.search(texto)
        fatos[chave] = m.group(0).strip() if m else None

    # frase negativa da secao 9
    m = re.search(r"^Durante o per[íi]odo.{0,60}?(?:interfer[êe]ncia|risco)s?[^.\n]*\.\s*$",
                  texto, re.IGNORECASE | re.MULTILINE)
    fatos["texto_s9"] = m.group(0).strip() if m else None

    # anexos (lista "Anexo N - titulo")
    fatos["anexos"] = re.findall(r"^Anexo\s+([IVXLC]+)\s*[–\-]\s*(.+)$", texto, re.MULTILINE)

    return fatos


# -----------------------------------------------------------------------------
# extracao dos fatos correspondentes do shadow (report_data + document_model)
# -----------------------------------------------------------------------------
_MAPA_SERVICO_NORM = None


def _lista_servicos_shadow(report_data):
    """Reconstroi a mesma regra de 1a ocorrencia (Q-9) usada pelo
    content_builder, para comparar o CONTEUDO da lista (nao a frase inteira,
    que tem verbo diferente do golden por desenho -- ver
    titulo-secao-51-padronizacao e comparison_rules servicos-sigla)."""
    from .content_builder import ServiceNamer
    namer = ServiceNamer()
    return namer.render_lista(report_data["services"])


def _data_br(iso):
    d = dt.date.fromisoformat(iso)
    return "%02d/%02d/%04d" % (d.day, d.month, d.year)


def montar_itens(report_data, document, golden_fatos, manifest):
    """Monta a lista de itens de comparacao (id, eixo, campo, golden, shadow,
    shadow_status, regra_id opcional)."""
    itens = []

    def add(id_, eixo, campo, golden, shadow, shadow_status=None, regra_id=None):
        itens.append({"id": id_, "eixo": eixo, "campo": campo, "golden": golden,
                      "shadow": shadow, "shadow_status": shadow_status, "regra_id": regra_id})

    road = report_data["road"]
    period = report_data["period"]
    contract = report_data["contract"]

    # ---------------------------------------------------------- metadata
    add("meta-report-number", "metadata", "capa.numero_relatorio",
        golden_fatos.get("relatorio"), "%03d" % report_data["report_number"])
    add("meta-rodovia", "metadata", "capa.rodovia", golden_fatos.get("rodovia"), road.get("rodovia"))
    add("meta-periodo", "metadata", "capa.periodo", golden_fatos.get("periodo"),
        "%s a %s" % (_data_br(period["start"]), _data_br(period["end"])))
    add("meta-executora", "metadata", "capa.empresa_executora",
        golden_fatos.get("executora"), road.get("empresa_executora_capa"))
    add("meta-supervisora", "metadata", "capa.empresa_supervisora",
        golden_fatos.get("supervisora"), road.get("empresa_supervisora"))
    add("meta-extensao", "metadata", "capa.extensao", golden_fatos.get("extensao"),
        road.get("extensao_texto"))

    # ---------------------------------------------------------- data (contratuais, §7)
    add("data-instrumento-convocacao", "data", "contract.instrumento_convocacao",
        golden_fatos.get("instrumento_convocacao"), contract.get("instrumento_convocacao"))
    add("data-contrato-obra", "data", "contract.contrato_obra",
        golden_fatos.get("contrato_obra"), contract.get("contrato_obra"))
    add("data-ordem-servico", "data", "contract.ordem_servico",
        golden_fatos.get("ordem_servico"), contract.get("ordem_servico"))
    add("data-ordem-servico-em", "data", "contract.ordem_servico_em",
        golden_fatos.get("ordem_servico_em"),
        _data_br(str(contract["ordem_servico_em"])) if contract.get("ordem_servico_em") else None)
    add("data-contrato-assinado-em", "data", "contract.contrato_assinado_em",
        golden_fatos.get("contrato_assinado_em"),
        _data_br(str(contract["contrato_assinado_em"])) if contract.get("contrato_assinado_em") else None,
        regra_id="contrato-assinado-em-nao-comprovado")

    # ---------------------------------------------------------- data (clima)
    for d in report_data["climate"]["by_day"]:
        dd_mmm = next((x["data_dd_mmm"] for x in period["days"] if x["date"] == d["date"]), d["date"])
        g = golden_fatos.get("climate_by_day", {}).get(dd_mmm)
        shadow_val = "-" if d.get("validation_status") == "HITL" else d["clima_label"]
        add("clima-%s" % d["date"], "data", "climate.by_day[%s]" % d["date"], g, shadow_val,
            shadow_status=d.get("validation_status"),
            regra_id=("clima-2026-09-01" if d["date"] == "2026-09-01"
                     else "clima-dia-sem-ficha" if shadow_val in ("-", "INDETERMINADO")
                     else None))

    # ---------------------------------------------------------- data (NC historico -- fato novo)
    hist = report_data["nonconformities"].get("history_by_status") or {}
    encerradas_shadow = hist.get("Encerrada")
    em_atraso_shadow = sum(v for k, v in hist.items() if k not in ("Encerrada", "Cancelada"))
    add("nc-encerradas", "data", "nonconformities.history_by_status.Encerrada",
        golden_fatos.get("nc_encerradas"), encerradas_shadow,
        regra_id="nc-historico-volatilidade")
    add("nc-em-atraso", "data", "nonconformities.history_by_status.em_atraso",
        golden_fatos.get("nc_em_atraso"), em_atraso_shadow,
        regra_id="nc-historico-volatilidade")

    # ---------------------------------------------------------- text (servicos, Q-9)
    lista_shadow = _lista_servicos_shadow(report_data)
    add("text-servicos", "text", "s1_apresentacao.servicos",
        golden_fatos.get("servicos_s1_frase"), lista_shadow,
        regra_id="servicos-sigla-equivalencia")

    # ---------------------------------------------------------- text (ambiental)
    env = report_data["environment"]
    add("text-ambiental-numero-licenca", "text", "s6_ambiental.numero_licenca",
        golden_fatos.get("licenca_numero"), None,
        shadow_status="HITL" if env.get("applies") else None,
        regra_id="ambiental-numero-licenca-nao-extraido")

    # ---------------------------------------------------------- text (interferencias, Q-10)
    add("text-s9-interferencias", "text", "s9_interferencias.texto",
        golden_fatos.get("texto_s9"),
        "Durante o período abrangido por este relatório, não foram registrados "
        "riscos ou interferências em campo.",
        shadow_status="HITL", regra_id="interferencias-sem-precedente-rd103")

    # ---------------------------------------------------------- section (titulos padronizados)
    add("section-titulo-5-1", "section", "s5_1_clima.titulo",
        golden_fatos.get("titulo_5_1"), "Condições climáticas do período",
        regra_id="titulo-secao-51-padronizacao")
    add("section-titulo-9", "section", "s9_interferencias.titulo",
        golden_fatos.get("titulo_9"), "RELATO DE RISCOS E INTERFERÊNCIAS EM CAMPO",
        # id renomeado na limpeza de regra obsoleta da Fase 3.2B (Q-12,
        # CHECKPOINT_3.2B_RENDERER_V2.md secao F) -- comparison_rules.yaml
        # passou a chamar "titulo-secao9-singular-plural"; esta referencia
        # tinha ficado com o nome antigo (KeyError em tests/test_comparison.py,
        # pre-existente ao inicio da Fase 3.3 -- corrigido aqui).
        regra_id="titulo-secao9-singular-plural")

    # ---------------------------------------------------------- structure (termo, anexos, figuras)
    add("structure-termo-npaginas", "text", "s10_termo_encerramento.n_paginas",
        golden_fatos.get("termo_n_paginas"), "___ (PLACEHOLDER)",
        shadow_status="PLACEHOLDER", regra_id="termo-encerramento-placeholder")

    anexos_golden = [normalizar_anexo_titulo(t) for _n, t in golden_fatos.get("anexos", [])]
    sec_anexos = next((s for s in document.sections if s.id == "s11_anexos"), None)
    itens_anexos_shadow = next((b.items for s in [sec_anexos] if s for b in s.blocks
                                if b.kind == "list"), []) if sec_anexos else []
    # cada item shadow ja vem formatado "Anexo I – <titulo>[ <numero>]" (ver
    # content_builder._secao11_anexos); extrai so o titulo apos o travessao.
    anexos_shadow = [normalizar_anexo_titulo(re.sub(r"^Anexo\s+[IVXLC]+\s*[–\-]\s*", "", it))
                     for it in itens_anexos_shadow]
    add("structure-anexos-ordem", "structure", "s11_anexos.ordem",
        anexos_golden, anexos_shadow)

    for id_, legenda in (
        ("fig-ficha-acompanhamento", "Figura 1 — Ficha de andamento físico da obra"),
        ("fig-avanco-grafico", "Gráfico 1 — Avanço global da obra"),
        ("fig-relatorio-fotografico", "Figuras 2..k — Relatório Fotográfico por serviço"),
        ("fig-ambiental", "Figuras 6..8 — painéis do Acompanhamento Ambiental"),
        ("fig-nc-historico", "Figura 9 + Gráficos 2/3 — histórico de NC"),
    ):
        add(id_, "structure", legenda, "imagem/gráfico real (golden)",
            "[RENDERIZAÇÃO PENDENTE]", shadow_status="PLACEHOLDER",
            regra_id="figura-placeholder")

    return itens


def normalizar_anexo_titulo(s):
    """Normaliza o titulo/id de um anexo para comparar golden x shadow
    ignorando diferencas de redacao ja conhecidas (GUIA_DE_REDACAO.md):
    'Parecer Técnico PA-NNN' == 'PA-NNN'; 'Acompanhamento de Saúde e
    Segurança do Trabalho' == 'Acompanhamento Saúde e Segurança do
    Trabalho'; ids internos (ex.: 'diagrama_unifilar') == titulo editorial."""
    n = _norm(s).replace("_", " ")
    n = re.sub(r"^parecer t[ée]cnico\s+", "", n)
    n = n.replace("acompanhamento de saude e seguranca do trabalho",
                 "acompanhamento saude e seguranca do trabalho")
    n = n.replace("acompanhamento sst", "acompanhamento saude e seguranca do trabalho")
    n = n.replace("nao conformidade", "nc")
    n = n.replace("diario de obras", "diario de obras")
    return n


# -----------------------------------------------------------------------------
# execucao completa
# -----------------------------------------------------------------------------
def comparar(report_data, document, manifest):
    golden = carregar_golden(manifest)
    if golden is None:
        return {"erro": "golden_reference.docx ausente ou não encontrado no manifest -- "
                        "comparação não pode ser feita", "itens": []}
    if golden.get("erro"):
        return {"erro": "falha ao ler o golden: %s" % golden["erro"], "itens": []}

    golden_fatos = extrair_fatos_golden(golden)
    regras = _carregar_regras()
    itens = montar_itens(report_data, document, golden_fatos, manifest)

    resultado = []
    for item in itens:
        classe, evidencia, regra_aplicada = classificar(item, regras)
        if classe not in CLASSES_VALIDAS:
            raise ValueError("classe invalida: %r (item %s)" % (classe, item["id"]))
        resultado.append(dict(item, classe=classe, evidencia=evidencia, regra_aplicada=regra_aplicada))

    contagem = {c: 0 for c in CLASSES_VALIDAS}
    for r in resultado:
        contagem[r["classe"]] += 1

    return {
        "gerado_em": dt.datetime.now().astimezone().isoformat(timespec="seconds"),
        "obra": report_data["metadata"]["road_key"],
        "periodo": "%s..%s" % (report_data["period"]["start"], report_data["period"]["end"]),
        "golden_path": (manifest.get("golden_reference", {}).get("docx", {}) or {}).get("path"),
        "eixos": EIXOS_VALIDOS,
        "classes": CLASSES_VALIDAS,
        "contagem_por_classe": contagem,
        "itens": resultado,
    }


def gerar_relatorio_md(comparison):
    linhas = []
    linhas.append("# Comparação Shadow × Golden — %s (%s)\n" % (
        comparison["obra"], comparison["periodo"]))
    linhas.append("Golden: `%s`\n" % comparison.get("golden_path"))
    linhas.append("\n## Contagem por classe\n")
    linhas.append("| Classe | Quantidade |")
    linhas.append("|---|---|")
    for c in comparison["classes"]:
        linhas.append("| %s | %d |" % (c, comparison["contagem_por_classe"].get(c, 0)))

    def _bloco(titulo, classes):
        linhas.append("\n## %s\n" % titulo)
        itens = [i for i in comparison["itens"] if i["classe"] in classes]
        if not itens:
            linhas.append("Nenhum item nesta categoria.\n")
            return
        for i in itens:
            linhas.append("- **%s** (`%s`, eixo %s) — golden=%r · shadow=%r%s" % (
                i["id"], i["campo"], i["eixo"], i["golden"], i["shadow"],
                (" — %s" % i["evidencia"]) if i.get("evidencia") else ""))

    linhas.append("\n---")
    _bloco("(a) Erros reais do Shadow (SHADOW_ERROR)", ("SHADOW_ERROR",))
    _bloco("(b) Diferenças corretas em relação ao golden", (
        "GOLDEN_DEFECT", "SEMANTICALLY_EQUIVALENT", "EXPECTED_DIFFERENCE", "SOURCE_LIMITATION"))
    _bloco("(c) Placeholders e pendências de confirmação humana", (
        "HITL_REQUIRED", "MISSING", "EXTRA"))
    return "\n".join(linhas) + "\n"
