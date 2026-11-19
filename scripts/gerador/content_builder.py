# -*- coding: utf-8 -*-
r"""
scripts/gerador/content_builder.py - Fase 3, Slice 3: report_data.json ->
document_model.Document (spec Sec.4/5).

NAO conhece caminho de arquivo nem planilha. Recebe so o dict `report_data`
(a saida de `gerador.canonical.construir()`) + os YAMLs de configuracao
(road_registry, report_versions, services_taxonomy, textos_v3) e monta o
modelo abstrato do documento. Zero IA: toda frase e regra deterministica
(DET), template parametrizado (TPL) ou extracao validada (VAL) -- nunca
geracao livre.

Le SEMPRE `validation_status` (prontidao para o documento), nunca `_status`
(resultado da extracao) -- sao conceitos distintos (spec Sec.1, caixa
ratificada em 2026-09-17).
"""
import datetime as dt

import unicodedata

from . import config as C
from . import document_model as M
from . import figuras_bridge as FB
from .photo_panel import PhotoPanelComposer, FotoRef

_PHOTO_COMPOSER = PhotoPanelComposer()


# -----------------------------------------------------------------------------
# Fase 3.3, Sec.3.1: sequencias de legenda INDEPENDENTES por tipo (Figura/
# Gráfico/Tabela), sempre GERADAS (nunca digitadas), sempre comecando em 1 --
# a lacuna do Checkpoint 3.2B era a Figura 1 nunca ser emitida, entao a
# sequencia de fotos comecava (corretamente numerada, mas sem a legenda
# anterior existir) em "Figura 2".
# -----------------------------------------------------------------------------
class CaptionSequencer:
    def __init__(self):
        self._contadores = {}

    def proximo(self, tipo):
        n = self._contadores.get(tipo, 0) + 1
        self._contadores[tipo] = n
        return n

    def legenda(self, tipo, chave_texto, status="READY", **kw):
        n = self.proximo(tipo)
        return M.caption(_fmt(chave_texto, n=n, **kw), status=status)


def _norm(s):
    s = unicodedata.normalize("NFKD", str(s or ""))
    s = "".join(c for c in s if not unicodedata.combining(c))
    return s.strip().lower()


def _fotos_do_servico(report_data, servico):
    """Contagem de fotos do servico no periodo inteiro, para o
    PhotoPanelComposer calcular o grid (P0.4). Casa `service_folder` (nome
    de pasta observado em `photos.by_day`) com o servico canonico por
    comparacao normalizada de texto -- a mesma LIMITATION ja registrada em
    report_schema.yaml (s5_2_fotografico.fotos.alerta): "a pasta de servico
    NAO e' classificacao confiavel". Nesta rodada isso so afeta o TAMANHO do
    grid (quantas fotos/celulas), nao decide integration_mode de anexo nem
    nenhuma outra regra de montagem -- por isso e' aceitavel como
    implementacao minima da interface do compositor."""
    alvo = _norm(servico.get("canonical_name"))
    alvo_chave = _norm(servico.get("key"))
    fotos = []
    for dia in report_data.get("photos", {}).get("by_day", []):
        if dia.get("status") != "OK":
            continue
        for s in dia.get("services", []):
            pasta = _norm(s.get("service_folder"))
            if pasta and (pasta in alvo or alvo in pasta or pasta in alvo_chave or alvo_chave in pasta):
                fotos.extend(FotoRef() for _ in range(s.get("n", 0)))
    return fotos

# -----------------------------------------------------------------------------
# numero por extenso (PT-BR) -- usado so pelo Termo de Encerramento (Sec.5)
# -----------------------------------------------------------------------------
_UNIDADES_M = ["", "um", "dois", "três", "quatro", "cinco", "seis", "sete", "oito", "nove"]
_UNIDADES_F = ["", "uma", "duas", "três", "quatro", "cinco", "seis", "sete", "oito", "nove"]
_DEZ_A_DEZENOVE = ["dez", "onze", "doze", "treze", "quatorze", "quinze", "dezesseis",
                   "dezessete", "dezoito", "dezenove"]
_DEZENAS = ["", "", "vinte", "trinta", "quarenta", "cinquenta", "sessenta", "setenta",
           "oitenta", "noventa"]
_CENTENAS = ["", "cento", "duzentos", "trezentos", "quatrocentos", "quinhentos",
            "seiscentos", "setecentos", "oitocentos", "novecentos"]


def por_extenso(n, feminino=False):
    """Numero cardinal por extenso em portugues, 0..999 (suficiente para a
    paginacao de um Relatorio Semanal). `feminino=True` para concordar com
    substantivos femininos (ex.: 'paginas' -> 'duas', nao 'dois')."""
    if not isinstance(n, int) or n < 0 or n > 999:
        raise ValueError("por_extenso: fora do intervalo suportado (0..999): %r" % (n,))
    unidades = _UNIDADES_F if feminino else _UNIDADES_M
    if n == 0:
        return "zero"
    if n == 100:
        return "cem"
    partes = []
    centena, resto = divmod(n, 100)
    if centena:
        partes.append(_CENTENAS[centena])
    if resto:
        if partes:
            partes.append("e")
        if resto < 10:
            partes.append(unidades[resto])
        elif resto < 20:
            partes.append(_DEZ_A_DEZENOVE[resto - 10])
        else:
            dezena, unidade = divmod(resto, 10)
            if unidade:
                partes.append("%s e %s" % (_DEZENAS[dezena], unidades[unidade]))
            else:
                partes.append(_DEZENAS[dezena])
    return " ".join(partes)


_MESES_EXTENSO = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho",
                  "agosto", "setembro", "outubro", "novembro", "dezembro"]


def _data_por_extenso(d):
    if isinstance(d, str):
        d = dt.date.fromisoformat(d)
    return "%d de %s de %d" % (d.day, _MESES_EXTENSO[d.month - 1], d.year)


def _data_br(iso):
    d = dt.date.fromisoformat(iso)
    return "%02d/%02d/%04d" % (d.day, d.month, d.year)


# -----------------------------------------------------------------------------
# Q-9: primeira ocorrencia completa "Nome (SIGLA)", ocorrencias seguintes so
# a sigla; se a taxonomia nao registra sigla, so o canonical_name sempre.
# -----------------------------------------------------------------------------
class ServiceNamer:
    def __init__(self):
        self._vistos = set()

    def render(self, servico):
        chave = servico["key"]
        nome = servico["canonical_name"]
        sigla = servico.get("sigla")
        if not sigla:
            return nome
        if chave in self._vistos:
            return sigla
        self._vistos.add(chave)
        return "%s (%s)" % (nome, sigla)

    def render_lista(self, servicos, minusculo=True):
        """Lista em prosa: 'a, b e c'. Aplica a regra de 1a ocorrencia na
        ORDEM da lista (a mesma que alimenta s1/s4/5.2 -- documento inteiro
        compartilha o mesmo rastreador quando construido por
        `construir_documento`). Servicos sem sigla entram em minusculo no
        corpo da frase (GUIA_DE_REDACAO.md §5); com sigla mantem a forma
        "Nome (SIGLA)"."""
        nomes = [self.render(s) for s in servicos]
        if minusculo:
            nomes = [n if (not n or "(" in n or n == n.upper()) else (n[0].lower() + n[1:])
                    for n in nomes]
        if not nomes:
            return ""
        if len(nomes) == 1:
            return nomes[0]
        return ", ".join(nomes[:-1]) + " e " + nomes[-1]


# -----------------------------------------------------------------------------
# helpers de texto/config
# -----------------------------------------------------------------------------
def _textos():
    return C._carregar("textos_v3.yaml")["textos"]


def _t(chave):
    ent = _textos()[chave]
    return ent.get("texto") or ent.get("template")


def _fmt(chave, **kw):
    return _t(chave).format(**kw)


def _fmt_n(chave, n, **kw):
    """Como `_fmt`, mas usa `template_singular` do catalogo quando n == 1
    (concordancia genérica; nunca por obra)."""
    ent = _textos()[chave]
    tpl = ent.get("template_singular") if n == 1 else None
    return (tpl if tpl else _t(chave)).format(n=n, **kw)


def _classe(chave):
    """Classe declarada no catalogo de textos ('A', 'LIMITATION', ['C','TPL']...).

    Devolve a classe como string quando ha uma so, para que o builder possa
    decidir o status do bloco pelo catalogo em vez de o cravar no codigo.
    """
    cl = _textos()[chave].get("classe")
    return cl if isinstance(cl, str) else (cl[0] if cl else None)


def _prov(chave):
    return [chave] if chave else []


# -----------------------------------------------------------------------------
# secoes
# -----------------------------------------------------------------------------
def _secao_capa(report_data, info):
    road = report_data["road"]
    period = report_data["period"]
    blocks = [
        M.paragraph(_t("capa_titulo"), status="READY"),
        M.cover_field("Rodovia", road.get("rodovia"), provenance=_prov("road.rodovia")),
        M.cover_field("Trecho", road.get("trecho"), provenance=_prov("road.trecho")),
        M.cover_field("Extensão", road.get("extensao_texto"), provenance=_prov("road.extensao_texto")),
        M.cover_field("Código SRE", ", ".join(road.get("codigo_sre") or []) or None,
                      provenance=_prov("road.codigo_sre")),
        M.cover_field("Lote", road.get("lote"), provenance=_prov("road.lote")),
        M.cover_field("Empresa executora", road.get("empresa_executora_capa"),
                      provenance=_prov("road.empresa_executora_capa")),
        M.cover_field("Empresa supervisora", road.get("empresa_supervisora"),
                      provenance=_prov("road.empresa_supervisora")),
        M.cover_field("Relatório", "%03d" % report_data["report_number"],
                      provenance=_prov("report_number")),
        M.cover_field("Período", "%s a %s" % (_data_br(period["start"]), _data_br(period["end"])),
                      provenance=_prov("period")),
    ]
    return M.Section(id="capa", number="", title="Capa", level=0,
                     page_setup="A4 retrato", page_profile="cover", blocks=blocks)


def _secao_indice(versao_secoes):
    itens = []
    for sec in versao_secoes:
        itens.append("%s — %s" % (sec["n"], sec["titulo"]))
        for sub in sec.get("subsecoes", []):
            itens.append("  %s — %s" % (sub["n"], sub["titulo"]))
    # as secoes 10 e 11 ja vêm de `versao_secoes` quando declaradas la; so
    # acrescenta o que faltar, para nao duplicar a entrada no indice.
    for n, titulo in (("10", "TERMO DE ENCERRAMENTO"), ("11", "ANEXOS")):
        if not any(i.lstrip().startswith(n + " —") for i in itens):
            itens.append("%s — %s" % (n, titulo))
    return M.Section(id="indice", number="", title="Índice", level=0,
                     page_setup="A4 retrato",
                     blocks=[M.heading("ÍNDICE", level=1), M.toc(itens)])


def _secao1_apresentacao(report_data, info, namer):
    road = report_data["road"]
    contract = report_data["contract"]
    period = report_data["period"]
    blocks = [M.heading("APRESENTAÇÃO", level=1)]

    blocks.append(M.paragraph(_fmt(
        "s1_contratual", contrato_prog=contract.get("contrato_prog") or "___",
        rodovia=road.get("rodovia"), trecho_extenso=road.get("trecho_extenso") or road.get("trecho"),
        extensao=road.get("extensao_texto")),
        # o status vem da propria entrada do catalogo: o texto literal foi
        # levantado do corpus (golden excluido) e deixou de ser LIMITATION.
        status="READY" if _classe("s1_contratual") == "A" else "LIMITATION",
        provenance=_prov("contract.contrato_prog") + _prov("road.rodovia"),
        note=None if _classe("s1_contratual") == "A" else
            "texto corrido literal do parágrafo contratual de abertura não capturado "
            "no corpus da descoberta (config/textos_v3.yaml s1_contratual)"))

    campos_contratacao = ["instrumento_convocacao", "contrato_obra", "contrato_assinado_em",
                          "ordem_servico", "ordem_servico_em"]
    faltando = [c for c in campos_contratacao if not contract.get(c)]
    txt_contratacao = (
        "A contratação da empresa executora {executora} deu-se por meio do Instrumento "
        "de Convocação Nº{instrumento_convocacao}, tendo sido firmado o Contrato "
        "Nº{contrato_obra}{assinado}. A Ordem de Serviço Nº{ordem_servico} foi emitida "
        "em {ordem_servico_em}.").format(
        executora=road.get("empresa_executora_contrato") or "___",
        instrumento_convocacao=contract.get("instrumento_convocacao") or "___",
        contrato_obra=contract.get("contrato_obra") or "___",
        assinado=(", assinado em %s" % _data_br(str(contract["contrato_assinado_em"]))
                 if contract.get("contrato_assinado_em") else ""),
        ordem_servico=contract.get("ordem_servico") or "___",
        ordem_servico_em=(_data_br(str(contract["ordem_servico_em"]))
                          if contract.get("ordem_servico_em") else "___"))
    blocks.append(M.paragraph(
        txt_contratacao, status="HITL" if faltando else "READY",
        provenance=["contract.%s" % c for c in campos_contratacao],
        note=("campos não comprovados por documento acessível a esta fatia: %s"
             % ", ".join(faltando)) if faltando else None))

    servicos_ordenados = report_data["services"]
    lista = namer.render_lista(servicos_ordenados)
    blocks.append(M.paragraph(_fmt(
        "s1_periodo_servicos", data_inicio=_data_br(period["start"]),
        data_fim=_data_br(period["end"]), lista_de_servicos=lista),
        status="READY", provenance=_prov("period") + _prov("services")))

    meetings = report_data.get("meetings") or []
    if meetings:
        m = meetings[0]
        blocks.append(M.paragraph(_fmt(
            "s1_reuniao_positiva", executora=road.get("empresa_executora_contrato") or "___",
            rodovia=road.get("rodovia"), data=m.get("date", "___")),
            status="READY", provenance=["meetings"]))
    else:
        blocks.append(M.paragraph(_fmt(
            "s1_reuniao_negativa", executora=road.get("empresa_executora_contrato") or "___",
            rodovia=road.get("rodovia")), status="READY", provenance=["meetings"]))

    return M.Section(id="s1_apresentacao", number="1", title="APRESENTAÇÃO", level=1,
                     page_setup="A4 retrato", blocks=blocks)


def _secao2_ficha(report_data, seq, figura_ficha=None):
    """`figura_ficha`: um FiguraResult (ou None) JA OBTIDO pelo chamador
    (shadow_report.cmd_candidate) via `figuras_bridge.tentar("ficha_acumulada",
    "gerar", caminho_ficha_acumulada=..., road_key=..., out_dir=...)` --
    content_builder.py continua sem conhecer caminho de arquivo nem planilha
    (regra da Sec.4 do docstring do modulo); quem resolve o caminho real e'
    sempre o pipeline. None (o default, usado pelos testes e por 'build'/
    'build2') -> mantem o placeholder de sempre."""
    blocks = [M.heading("FICHA DE ACOMPANHAMENTO", level=1)]
    if FB.disponivel(figura_ficha):
        blocks.append(M.figure(
            FB.campo(figura_ficha, "path"), FB.campo(figura_ficha, "largura_mm"),
            FB.campo(figura_ficha, "altura_mm"),
            status=FB.mapear_status(FB.campo(figura_ficha, "status")),
            note=FB.campo(figura_ficha, "note"), provenance=["figuras.ficha_acumulada"]))
    else:
        blocks.append(M.figure_placeholder(
            "[FIGURA 1 — FICHA DE CAMPO ACUMULADA — RENDERIZAÇÃO PENDENTE]",
            note="fonte: Ficha de Campo Acumulada -- scripts/gerador/figuras/ficha_acumulada.py "
                "ainda não disponível/aplicável nesta rodada (Frente A, ver figuras_bridge.py)"))
    # Sec.3.1: fecha a lacuna do Checkpoint 3.2B -- a legenda de Figura 1 é
    # SEMPRE emitida, exista ou não a imagem real ainda.
    blocks.append(seq.legenda("Figura", "s2_legenda_figura1"))
    return M.Section(id="s2_ficha_acompanhamento", number="2", title="FICHA DE ACOMPANHAMENTO",
                     level=1, page_setup="A4 retrato", blocks=blocks)


def _fmt_pct(v):
    if v is None:
        return None
    return ("%.2f %%" % (v * 100 if v <= 1 else v)).replace(".", ",")


def _secao3_avanco(report_data, seq, figura_avanco=None, info=None):
    """`figura_avanco`: FiguraResult (ou None) já obtido pelo chamador via
    `figuras_bridge.tentar("avanco_grafico", "gerar", report_data=..., out_dir=...)`
    -- mesma regra de _secao2_ficha (content_builder não abre arquivo).

    Fase 3.4 (v0.2), GF-1 (TEXTO-01): os 3 parágrafos-conceito (avanço global /
    avanço concluído / avanço concluído e atacado) são texto-padrão levantado
    do corpus (config/textos_v3.yaml s3_avanco_*, 103/103 = 100%, golden
    excluído) -- nunca dependem de dado da semana, por isso são sempre READY."""
    progress = report_data["progress"]
    concluido = _fmt_pct(progress.get("global_concluido"))
    atacado = _fmt_pct(progress.get("global_concluido_atacado"))
    status = "READY" if concluido and atacado else "HITL"
    blocks = [
        M.heading("AVANÇO GLOBAL DA OBRA", level=1),
        M.paragraph(_t("s3_avanco_conceito")),
        M.paragraph(_t("s3_avanco_concluido_conceito")),
        M.paragraph(_t("s3_avanco_concluido_atacado_conceito")),
    ]
    # Fase 3.4 (v0.3), TEXTO-01: a frase-ponte que apresenta o Gráfico 1 é
    # VARIAÇÃO POR OBRA, não texto universal -- RD-101 e RD-103 a emitem
    # (23/26 e 22/25 no corpus), RD-102A e RD-104 não (6/26 cada). O
    # liga/desliga vem do road_registry; texto único, sem número dentro.
    if (info or {}).get("s3_frase_ponte_grafico"):
        blocks.append(M.paragraph(_t("s3_frase_ponte_grafico")))
    blocks += [
        M.paragraph(_fmt("s3_avanco_texto", concluido=concluido or "___",
                         concluido_atacado=atacado or "___"),
                    status=status, provenance=_prov("progress.global_concluido")
                    + _prov("progress.global_concluido_atacado")),
    ]
    if FB.disponivel(figura_avanco):
        blocks.append(M.figure(
            FB.campo(figura_avanco, "path"), FB.campo(figura_avanco, "largura_mm"),
            FB.campo(figura_avanco, "altura_mm"),
            status=FB.mapear_status(FB.campo(figura_avanco, "status")),
            note=FB.campo(figura_avanco, "note"), provenance=["figuras.avanco_grafico"]))
    else:
        blocks.append(M.figure_placeholder("[GRÁFICO 1 — AVANÇO FÍSICO — RENDERIZAÇÃO PENDENTE]"))
    blocks.append(seq.legenda("Gráfico", "s3_legenda_grafico1"))
    return M.Section(id="s3_avanco_global", number="3", title="AVANÇO GLOBAL DA OBRA",
                     level=1, page_setup="A4 retrato", blocks=blocks)


def _secao4_unifilar(report_data, info, namer, titulo):
    road = report_data["road"]
    blocks = [
        M.heading(titulo, level=1),
        M.paragraph(_fmt("s4_unifilar_texto_base", rodovia=road.get("rodovia")),
                    provenance=_prov("road.rodovia")),
        M.paragraph(_fmt("s4_unifilar_frase_servicos",
                         lista_de_servicos=namer.render_lista(report_data["services"])),
                    provenance=_prov("services")),
    ]
    if info.get("tem_tabela_bueiros"):
        blocks.append(M.table(
            headers=["Estaca", "Tipo", "Situação"], rows=[],
            status="PLACEHOLDER",
            note="Tabela 1 - Bueiros projetados (extração da tabela dedicada da Ficha de "
                "Campo fora do escopo desta fatia)"))
    return M.Section(id="s4_unifilar", number="4", title=titulo, level=1,
                     page_setup="A4 retrato", blocks=blocks)


def _secao5_1_clima(report_data, info, titulo_secao, seq):
    by_day = report_data["climate"]["by_day"]
    dd_mmm_por_data = {d["date"]: d["data_dd_mmm"] for d in report_data["period"]["days"]}
    linha_data = [dd_mmm_por_data.get(d["date"], d["date"]) for d in by_day]
    linha_dia = [d["label"] for d in by_day]
    linha_clima = []
    algum_hitl = False
    for d in by_day:
        if d.get("validation_status") == "HITL":
            algum_hitl = True
            linha_clima.append("-")  # T13: grafia do acervo p/ dia nao comprovado
        else:
            linha_clima.append(d["clima_label"])
    # tabela de 3 LINHAS (DATA/DIA/CLIMA) x 7 dias-COLUNA (report_schema.yaml
    # s5_1_clima.tabela: "linhas: [DATA, DIA, CLIMA], colunas: 7") -- a
    # primeira celula de cada linha e' o rotulo, nao um cabecalho de coluna.
    rows = [["DATA"] + linha_data, ["DIA"] + linha_dia, ["CLIMA"] + linha_clima]
    headers = None
    status_tabela = "HITL" if algum_hitl else "READY"
    note = None
    if algum_hitl:
        dias_hitl = [d["date"] for d in by_day if d.get("validation_status") == "HITL"]
        note = ("dia(s) sem Ficha de Campo e sem fonte alternativa comprovada: %s "
                "-- celula renderizada '-' (T13_CLIMA_DIA_SEM_FICHA.md); NAO e default "
                "nem repeticao do dia anterior" % ", ".join(dias_hitl))

    blocks = [
        M.heading(titulo_secao, level=2),
        # Fase 3.4 (v0.2), GF-1 (TEXTO-02): parágrafo intro + as 3 definições
        # climáticas -- texto-padrão (config/textos_v3.yaml s5_1_*), nunca
        # dado da semana, sempre READY.
        M.paragraph(_t("s5_1_intro_definicoes")),
        M.list_block([_t("s5_1_definicao_bom"), _t("s5_1_definicao_chuvoso"),
                     _t("s5_1_definicao_impraticavel")]),
        M.paragraph(_t("s5_1_abertura_tabela")),
        # role="climate_table" (P1.9): rotulo semantico para o renderer v2
        # aplicar as 8 larguras reais / bordas 0,5pt / shading F2F2F2 sem
        # depender de posicao no documento -- e' a MESMA tabela sempre
        # (DATA/DIA/CLIMA x 7 dias), reconhecida pelo papel dela, nao pelo
        # numero da secao.
        M.table(headers, rows, status=status_tabela,
               provenance=["climate.by_day.%s" % d["date"] for d in by_day], note=note,
               role="climate_table"),
    ]
    # Sec.3.1: legenda de Tabela 1 -- fecha a 3a lacuna do Checkpoint 3.2B.
    blocks.append(seq.legenda("Tabela", "s5_1_legenda_tabela1"))
    dias_sem = report_data["climate"]["days_without_work"]
    if dias_sem:
        datas_br = [_data_br(d) for d in dias_sem]
        if len(datas_br) == 1:
            blocks.append(M.paragraph(_fmt("s5_1_frase_sem_execucao_um_dia", data=datas_br[0]),
                                      provenance=_prov("climate.days_without_work")))
        else:
            blocks.append(M.paragraph(
                _fmt("s5_1_frase_sem_execucao_varios_dias",
                    datas=", ".join(datas_br[:-1]) + " e " + datas_br[-1]),
                provenance=_prov("climate.days_without_work")))
    return blocks


def _secao5_2_fotografico(report_data, namer, seq, paineis_fotograficos=None):
    """`paineis_fotograficos`: dict {servico_key: FiguraResult} já obtido pelo
    chamador via `figuras_bridge.tentar("painel_fotografico", "montar",
    report_data=..., manifest=..., out_dir=...)` -- o "item central" da Fase
    3.3 (spec Sec.2.3). None/ausente para um serviço -> mantém o placeholder
    do PhotoPanelComposer (P0.4, Checkpoint 3.2B), nunca foto a foto."""
    road = report_data["road"]
    paineis_fotograficos = paineis_fotograficos or {}
    blocks = []
    for servico in report_data["services"]:
        nome_sub = servico["canonical_name"]
        blocks.append(M.heading(nome_sub, level=3))
        # Sec.3.1 (Fase 3.3): numeracao de Figura vem do sequenciador
        # compartilhado do documento inteiro -- Figura 1 e' a ficha
        # acumulada (secao 2), entao aqui ja comeca corretamente em 2 sem
        # nenhum contador local proprio.
        n_fig = seq.proximo("Figura")
        legenda = _fmt("s5_2_legenda_foto", n=n_fig, servico_canonico=servico["canonical_name"],
                       rodovia=road.get("rodovia"))
        resultado_painel = paineis_fotograficos.get(servico["key"])
        if FB.disponivel(resultado_painel):
            blocks.append(M.figure(
                FB.campo(resultado_painel, "path"), FB.campo(resultado_painel, "largura_mm"),
                FB.campo(resultado_painel, "altura_mm"),
                status=FB.mapear_status(FB.campo(resultado_painel, "status")),
                note=FB.campo(resultado_painel, "note"),
                provenance=["figuras.painel_fotografico.%s" % servico["key"]]))
        else:
            # P0.4: NUNCA inserir foto por foto -- toda foto do servico passa
            # pelo PhotoPanelComposer, que decide o grid (3 colunas fixas,
            # linhas = ceil(n/3)) e devolve o painel composto (ou, nesta
            # rodada sem a figura real, um placeholder correto pelo
            # componente certo -- ver photo_panel.py).
            fotos = _fotos_do_servico(report_data, servico)
            painel = _PHOTO_COMPOSER.compose(nome_sub, fotos)
            blocks.append(M.figure_placeholder(
                "[%s — RENDERIZAÇÃO PENDENTE]" % legenda,  # legenda ja' comeca com "Figura {n} -- ..."
                status="LIMITATION" if painel.status == "LIMITATION" else "PLACEHOLDER",
                provenance=["photos.by_day"],
                note=painel.note))
        blocks.append(M.caption(legenda))
    return blocks


def _secao5_execucao(report_data, info, seq, paineis_fotograficos=None):
    titulo_51 = "Condições climáticas do período"
    blocks_51 = _secao5_1_clima(report_data, info, titulo_51, seq)
    namer_local = ServiceNamer()  # 5.2 usa titulo completo por subsecao (cada subtitulo e' seu proprio "1a ocorrencia" visual)
    blocks_52 = _secao5_2_fotografico(report_data, namer_local, seq, paineis_fotograficos)
    # Fase 3.4 (v0.2), item 4 do plano (política de quebra de página): o
    # golden agrupa §4+§5+5.1+"5.2 intro" na MESMA página A4 (p6) -- só o
    # painel fotográfico de verdade abre a página A3. O heading "Relatório
    # Fotográfico" + o parágrafo fixo (s5_2_intro) por isso ficam na SEÇÃO A4
    # (sec), não na seção A3 (sec_52) -- TOC não é afetado (o campo TOC
    # escaneia por ESTILO Heading em todo o documento, não por Section).
    blocks_52_intro = [M.heading("Relatório Fotográfico", level=2), M.paragraph(_t("s5_2_intro"))]
    sec = M.Section(id="s5_execucao_obra", number="5", title="EXECUÇÃO DE OBRA", level=1,
                    page_setup="A4 retrato", blocks=[M.heading("EXECUÇÃO DE OBRA", level=1)]
                    + blocks_51 + blocks_52_intro)
    sec_52 = M.Section(id="s5_2_fotografico", number="5.2", title="Relatório Fotográfico",
                       level=2, page_setup="A3 retrato", page_profile="photo", blocks=blocks_52)
    return [sec, sec_52]


def _pa_numero(report_data):
    for a in report_data.get("annexes") or []:
        if a.get("id") == "parecer_tecnico":
            return a.get("pa_numero")
    return None


def _licencas_kwargs(modelo, resultado):
    """Monta os kwargs do template s6_licencas_modelo_<modelo> (config/
    textos_v3.yaml) a partir do FiguraResult de
    figuras.ambiental.gerar_indices(...) -- provenance.road_specific_metrics
    traz total_licencas/emitidas/pendentes (Fase 3.4, DADO-07/08, corrigido
    pela outra frente desta rodada no filtro por Rodovia). Só o modelo C
    (RD-103, o aplicável a esta rodada) tem TODOS os campos comprovados por
    essa fonte com o MESMO nome semântico do template; os demais (A/B/D)
    ficam HITL -- "válidas"/"No prazo" não é uma contagem própria do extrator
    hoje (só 'Vigente'/'Aguardando a liberação'), então mapear seria inventar
    equivalência não comprovada (GUIA_DE_REDACAO.md §6: nunca inferir dado)."""
    prov = FB.campo(resultado, "provenance") or {}
    m = prov.get("road_specific_metrics") or {}
    emitidas, pendentes, total = m.get("emitidas"), m.get("pendentes"), m.get("total_licencas")
    if modelo == "C" and emitidas is not None and pendentes is not None:
        return {"n_emitidas": emitidas, "n_pendentes": pendentes}, "READY"
    if modelo == "A" and total is not None:
        return {"n_validas": total}, "HITL"
    if modelo == "B" and None not in (total, emitidas, pendentes):
        return {"n_solicitadas": total, "n_emitidas": emitidas, "n_pendentes": pendentes}, "HITL"
    if modelo == "D" and None not in (total, pendentes):
        return {"n_solicitadas": total, "n_no_prazo": emitidas, "n_pendentes": pendentes}, "HITL"
    return None, "LIMITATION"


def _condicionantes_kwargs(resultado):
    """Idem `_licencas_kwargs`, para figuras.ambiental.gerar_condicionantes(...)
    -- provenance.n_condicionantes (total) e
    provenance.road_specific_metrics.contagem_por_status (por status, domínio
    em ambiental.py:_ORDEM_STATUS_CONDICIONANTE). n_atencao soma "Pendente" +
    "Próximo do vencimento" (DADO-09: a 3ª categoria só existe depois da
    Frente B terminar a correção do extrator; até lá soma o que a chave
    trouxer, sem inventar)."""
    prov = FB.campo(resultado, "provenance") or {}
    m = prov.get("road_specific_metrics") or {}
    contagem = m.get("contagem_por_status")
    if contagem is None:
        return None, "LIMITATION"
    n_atencao = (contagem.get("Pendente", 0) or 0) + (contagem.get("Próximo do vencimento", 0) or 0)
    return n_atencao, "READY"


def _nota_historica(id_dado):
    """Texto da nota que acompanha um dado suprimido por ser irreconstituível."""
    m = C.motivo_historico(id_dado) or {}
    return ("modo HISTORICAL_SHADOW: afirmação omitida porque a fonte guarda só o "
            "estado atual, não o da semana do relatório. %s Evidência: %s"
            % (m.get("motivo", ""), m.get("evidencia", "")))


def _omitir_por_ser_historico(modo, id_dado):
    """True quando este dado NÃO pode ser afirmado nesta geração.

    A decisão sai inteira de config/generation_modes.yaml: aqui não se julga
    o que é volátil, só se aplica o que está declarado lá com evidência.
    """
    return (modo == C.MODO_HISTORICAL_SHADOW
            and C.acao_historica(id_dado) == "omitir_afirmacao")


def _secao6_ambiental(report_data, info, seq, figuras_ambientais=None,
                      modo=C.MODO_CURRENT):
    """`figuras_ambientais`: dict OPCIONAL {"indices"|"condicionantes"|"qrcode":
    FiguraResult|None} já obtido pelo chamador via `figuras_bridge.tentar(
    "ambiental", "gerar_indices"/"gerar_condicionantes"/"gerar_qrcode", ...)`
    -- mesma regra das demais figuras (content_builder não abre planilha).
    None/ausente para uma chave -> mantém o placeholder de sempre (Correção 3,
    Fase 3.3).

    Fase 3.4 (v0.2), GF-1 (TEXTO-05): 4 parágrafos descritivos que faltavam
    (monitoramento quinzenal, Dashboard Ambiental, contagem de licenças,
    condicionantes/Figura 7) -- os 2 primeiros são texto-padrão fixo; os 2
    últimos são template escolhido por road_registry.<obra> (nunca
    condicional por nome de obra no código), preenchidos com o dado real quando a figura correspondente
    já veio pronta da Frente B (senão ficam ausentes -- nunca um número
    inventado)."""
    env = report_data["environment"]
    figuras_ambientais = figuras_ambientais or {}
    blocks = [M.heading("ACOMPANHAMENTO AMBIENTAL", level=1)]
    if not env.get("applies"):
        blocks.append(M.paragraph(
            "Não se aplica ao período abrangido por este relatório o ciclo quinzenal de "
            "acompanhamento ambiental.", status="READY", provenance=_prov("environment.applies")))
    else:
        pa_numero = _pa_numero(report_data)
        blocks.append(M.paragraph(
            _fmt("s6_intro_monitoramento_quinzenal", pa_numero=pa_numero or "___"),
            status="READY" if pa_numero else "HITL",
            provenance=_prov("environment.applies") + _prov("annexes.parecer_tecnico"),
            note=None if pa_numero else "Parecer Técnico do ciclo quinzenal não identificado nesta fatia"))
        blocks.append(M.paragraph(_t("s6_dashboard_ambiental")))
        blocks.append(M.paragraph(_t("s6_texto_sem_numero").format(rodovia=report_data["road"]["rodovia"]),
                                  status="HITL", provenance=_prov("environment.applies"),
                                  note="número da Licença Ambiental Única não extraído nesta fatia"))

        modelo_licencas = info.get("ambiental_modelo_licencas")
        resultado_indices = figuras_ambientais.get("indices")
        if _omitir_por_ser_historico(modo, "environment.licencas_kpis"):
            # v0.3: a frase "possui N emitidas e M pendentes" seria uma
            # afirmação sobre o acervo da semana 49 feita com o acervo de
            # hoje. Não se emite, e não se copia do golden -- o leitor fica
            # sabendo que o dado existiu e por que não foi reconstruído.
            blocks.append(M.paragraph(
                "A composição de licenças emitidas e pendentes na data de referência "
                "deste relatório não foi reconstruída.",
                status="LIMITATION",
                provenance=_prov("environment.licencas") + ["config/generation_modes.yaml"],
                note=_nota_historica("environment.licencas_kpis")))
        elif modelo_licencas and FB.disponivel(resultado_indices):
            kwargs, status_lic = _licencas_kwargs(modelo_licencas, resultado_indices)
            if kwargs is not None:
                blocks.append(M.paragraph(
                    _fmt("s6_licencas_modelo_%s" % modelo_licencas.lower(), **kwargs),
                    status=status_lic, provenance=["figuras.ambiental.indices"],
                    note=None if status_lic == "READY" else
                        "mapeamento do modelo de frase por obra ainda não confirmado 1:1 com o "
                        "domínio de status do extrator (ver docstring de _licencas_kwargs)"))

        for chave, legenda in (("indices", "Índices de Acompanhamento Ambiental"),
                               ("condicionantes", "Status de Condicionantes Ambientais"),
                               ("qrcode", "Gestão Ambiental")):
            # Sec.3.1 (Fase 3.3): mesmo sequenciador do documento inteiro --
            # nunca mais o literal "n" no rotulo (bug corrigido junto com a
            # lacuna das 3 legendas do Checkpoint 3.2B).
            n = seq.proximo("Figura")
            resultado = figuras_ambientais.get(chave)
            if FB.disponivel(resultado):
                blocks.append(M.figure(
                    FB.campo(resultado, "path"), FB.campo(resultado, "largura_mm"),
                    FB.campo(resultado, "altura_mm"),
                    status=FB.mapear_status(FB.campo(resultado, "status")),
                    note=FB.campo(resultado, "note"),
                    label="[FIGURA %d — %s — IMAGEM AUSENTE]" % (n, legenda),
                    provenance=["figuras.ambiental.%s" % chave]))
            else:
                blocks.append(M.figure_placeholder(
                    "[FIGURA %d — %s — RENDERIZAÇÃO PENDENTE]" % (n, legenda)))
            # a frase-ponte das condicionantes (Figura 7 do golden) vem LOGO
            # DEPOIS da própria figura de condicionantes (mesma ordem do
            # corpus: painel -> frase explicando o painel).
            if chave == "condicionantes":
                conector = info.get("ambiental_conector_condicionantes")
                preposicao = info.get("ambiental_condicionantes_preposicao_vencimento")
                if _omitir_por_ser_historico(modo, "environment.condicionantes_status"):
                    # v0.3: o parágrafo descritivo do painel PERMANECE (ele não
                    # afirma quantidade); some só a oração final, que afirmaria
                    # "há N próximas do vencimento" -- número que oscilou de 31
                    # para 2 em dez dias sobre as mesmas 45 condicionantes.
                    blocks.append(M.paragraph(
                        _fmt("s6_condicionantes_generica", n=n),
                        provenance=["figuras.ambiental.condicionantes"]))
                    blocks.append(M.paragraph(
                        "A quantidade de condicionantes próximas do vencimento na data de "
                        "referência deste relatório não foi reconstruída.",
                        status="LIMITATION",
                        provenance=["figuras.ambiental.condicionantes",
                                    "config/generation_modes.yaml"],
                        note=_nota_historica("environment.condicionantes_status")))
                elif conector and preposicao and FB.disponivel(resultado):
                    n_atencao, status_cond = _condicionantes_kwargs(resultado)
                    if status_cond == "READY":
                        if n_atencao > 0:
                            blocks.append(M.paragraph(_fmt(
                                "s6_condicionantes_positiva", n=n, conector=conector,
                                preposicao=preposicao, n_atencao=n_atencao),
                                provenance=["figuras.ambiental.condicionantes"]))
                        else:
                            blocks.append(M.paragraph(_fmt(
                                "s6_condicionantes_negativa", n=n, conector=conector,
                                preposicao=preposicao),
                                provenance=["figuras.ambiental.condicionantes"]))
    if env.get("kpis_status") == "MVP_LIMITATION":
        blocks.append(M.paragraph(
            "Os indicadores consolidados do painel ambiental (índices, condicionantes e "
            "gestão) não foram reconstruídos nesta fatia da automação.",
            status="LIMITATION", note=env.get("kpis_nota")))
    return M.Section(id="s6_ambiental", number="6", title="ACOMPANHAMENTO AMBIENTAL", level=1,
                     page_setup="A4 retrato", blocks=blocks)


def _secao7_diario(report_data, info):
    """Fase 3.4 (v0.2), GF-1 (TEXTO-03): faltavam 3 das 4 frases do golden --
    (1) histórico físico->digital, (3) formalização/assinatura, (4)
    "encontra-se anexo" -- só (2) já existia (s7_texto_base). (1) e (4) usam
    FATO por obra (road_registry), nunca hardcode da RD-103 (GUIA_DE_REDACAO.md
    §7: a data de migração NÃO é a mesma nas 4 obras; a vírgula antes de
    "encontra-se" também é forma consistente por obra, não erro de digitação
    isolado -- ver nota em textos_v3.yaml)."""
    blocks = [M.heading("DIÁRIO DE OBRAS", level=1)]

    data_fim_fisico = info.get("diario_migracao_fisico_ate")
    data_inicio_digital = info.get("diario_migracao_digital_desde")
    if data_fim_fisico and data_inicio_digital:
        blocks.append(M.paragraph(_fmt(
            "s7_historico_formato",
            data_fim_fisico=_data_por_extenso(data_fim_fisico),
            data_inicio_digital=_data_por_extenso(data_inicio_digital)),
            provenance=["road.diario_migracao_fisico_ate", "road.diario_migracao_digital_desde"]))

    blocks.append(M.paragraph(_t("s7_texto_base")))
    blocks.append(M.paragraph(_t("s7_formalizacao_assinatura")))

    chave_anexo = ("s7_diario_anexo_com_virgula" if info.get("diario_anexo_com_virgula")
                  else "s7_diario_anexo_sem_virgula")
    blocks.append(M.paragraph(_t(chave_anexo)))

    return M.Section(id="s7_diario_obras", number="7", title="DIÁRIO DE OBRAS", level=1,
                     page_setup="A4 retrato", blocks=blocks)


# Fase 3.4 (v0.2), GF-1 (TEXTO-04): ordem e rótulos da frase-ponte do Gráfico
# 2 -- "Aberta" primeiro, "Encerrada" por último, sempre (padrão observado em
# 23/27 variantes do corpus); só entram categorias com contagem > 0.
_ORDEM_STATUS_GRAFICO2 = ["Aberta", "Aguardando Projeto Executivo", "Cancelada",
                         "Em Análise", "Aguardando Evidências", "Encerrada"]
_LABEL_STATUS_GRAFICO2 = {
    "Aberta": "em aberto", "Aguardando Projeto Executivo": "aguardando projeto executivo",
    "Cancelada": "canceladas", "Em Análise": "em análise",
    "Aguardando Evidências": "aguardando evidências", "Encerrada": "encerradas",
}


def _distribuicao_status_prosa(hist):
    """'2 em aberto, 6 aguardando projeto executivo, ... e 22 encerradas' --
    prosa em PT-BR (vírgula + 'e' antes do último) a partir de
    nonconformities.history_by_status (mesma chave já usada pela tabela).
    Devolve None se `hist` não tiver nenhuma categoria com contagem > 0
    (nunca inventa número)."""
    clausulas = ["%d %s" % (hist[st], _LABEL_STATUS_GRAFICO2[st])
                for st in _ORDEM_STATUS_GRAFICO2 if hist.get(st)]
    if not clausulas:
        return None
    if len(clausulas) == 1:
        return clausulas[0]
    return ", ".join(clausulas[:-1]) + " e " + clausulas[-1]


def _secao8_nc(report_data, info, seq, figuras_nc=None):
    """`figuras_nc`: dict OPCIONAL {"notificacoes"|"grafico_status"|
    "grafico_natureza": FiguraResult|None} já obtido pelo chamador via
    `figuras_bridge.tentar("nc_graficos", "gerar_figura_notificacoes"/
    "gerar_grafico_status"/"gerar_grafico_natureza", ...)` -- mesma regra das
    demais figuras. None/ausente para uma chave -> mantém o placeholder de
    sempre (Correção 3, Fase 3.3)."""
    figuras_nc = figuras_nc or {}
    nc = report_data["nonconformities"]
    blocks = [M.heading("ACOMPANHAMENTO DE NÃO CONFORMIDADES", level=1)]

    blocks.append(M.heading("Não conformidades emitidas pelo Programa no período", level=2))
    issued = nc.get("issued_in_period") or []
    if not issued:
        blocks.append(M.paragraph(_t("s8_1_negativa"), provenance=_prov("nonconformities.issued_in_period")))
    else:
        blocks.append(M.paragraph(_fmt_n("s8_1_positiva_intro", len(issued))))
        itens = [_fmt("s8_1_item", sentence=i["sentence"]) for i in issued]
        blocks.append(M.list_block(itens, provenance=_prov("nonconformities.issued_in_period")))

    blocks.append(M.heading("Não conformidades emitidas pela Órgão Contratante no período", level=2))
    blocks.append(M.paragraph(_t("s8_2_negativa"), provenance=_prov("nonconformities.orgao")))

    blocks.append(M.heading("Histórico de não conformidades emitidas e não concluídas", level=2))
    hist = nc.get("history_by_status")
    if hist:
        rows = [[k, str(v)] for k, v in hist.items()]
        recon = nc.get("historical_reconstruction", {})
        vol = recon.get("historical_nc_state") == "UNRECONSTRUCTABLE_FROM_CURRENT_SOURCE"
        blocks.append(M.table(
            ["Status", "Quantidade"], rows,
            status="LIMITATION" if vol else "READY",
            provenance=_prov("nonconformities.history_by_status"),
            note=("contagem reflete o estado ATUAL da planilha (TODAY()), não o estado "
                 "da semana do relatório -- %s" % recon.get("motivo", "")) if vol else None,
            # Sec.3.3 (Fase 3.3): mesmo tratamento visual da tabela do clima
            # (bordas 0,5pt, cabeçalho sombreado F2F2F2, Arial) -- rótulo
            # semântico para o renderer v2 reconhecer sem depender de posição.
            role="nc_status_table"))
        # Sec.3.1: mesmo sequenciador do documento inteiro -- a primeira
        # legenda (painel de notificações) e' Figura; as duas seguintes
        # (status/natureza) sao Gráfico, igual ao GOLDEN_VISUAL_SPEC §6
        # (Figura 9 / Gráfico 2 / Gráfico 3).
        for tipo, chave, legenda in (("Figura", "notificacoes", "Acompanhamento de notificações"),
                                     ("Gráfico", "grafico_status", "Distribuição de NC por status"),
                                     ("Gráfico", "grafico_natureza", "Distribuição por natureza")):
            n = seq.proximo(tipo)
            # GF-1: as frases-ponte vêm ANTES da figura que introduzem (mesma
            # ordem do corpus: "O Gráfico N apresenta..." sempre precede o
            # próprio gráfico).
            if chave == "grafico_status":
                distrib = _distribuicao_status_prosa(hist)
                if distrib:
                    blocks.append(M.paragraph(
                        _fmt("s8_3_grafico2_ponte", n=n, distribuicao_status=distrib),
                        provenance=_prov("nonconformities.history_by_status")))
            elif chave == "grafico_natureza":
                chave_txt = ("s8_3_grafico3_ponte_a_seguir"
                            if info.get("nc_grafico3_estilo_referencia") == "a_seguir"
                            else "s8_3_grafico3_ponte_numero")
                blocks.append(M.paragraph(_fmt(chave_txt, n=n)))
            resultado = figuras_nc.get(chave)
            if FB.disponivel(resultado):
                blocks.append(M.figure(
                    FB.campo(resultado, "path"), FB.campo(resultado, "largura_mm"),
                    FB.campo(resultado, "altura_mm"),
                    status=FB.mapear_status(FB.campo(resultado, "status")),
                    note=FB.campo(resultado, "note"),
                    label="[%s %d — %s — IMAGEM AUSENTE]" % (tipo.upper(), n, legenda),
                    provenance=["figuras.nc.%s" % chave]))
            else:
                blocks.append(M.figure_placeholder(
                    "[%s %d — %s — RENDERIZAÇÃO PENDENTE]" % (tipo.upper(), n, legenda)))
            # GF-1 (TEXTO-04): parágrafo de definição de status, logo depois
            # do painel de notificações (mesma posição do corpus) e antes da
            # frase-ponte do Gráfico 2.
            if chave == "notificacoes":
                blocks.append(M.paragraph(_t("s8_3_status_definicao")))
    else:
        blocks.append(M.paragraph("Histórico de não conformidades não disponível nesta fatia.",
                                  status="LIMITATION"))
    return M.Section(id="s8_nao_conformidades", number="8",
                     title="ACOMPANHAMENTO DE NÃO CONFORMIDADES", level=1,
                     page_setup="A4 retrato", blocks=blocks)


def _secao9_interferencias(report_data, info):
    """Q-10 (spec Sec.1): sem ocorrencia -> gera a frase negativa, classe TPL,
    validation_status HITL (nao bloqueia a montagem).

    P1.11 (Q-12, FASE3_2B_RENDERER_V2_SPEC.md): o titulo vem SEMPRE de
    `road_registry.titulo_secao_9`, por obra -- nunca literal no codigo. As 4
    obras ja divergem entre si em singular/plural e nenhuma delas leva
    "RISCOS" alem do que a propria obra registra (GUIA_DE_REDACAO.md R-05);
    nao acrescentar a palavra aqui alem do que o registro ja contem."""
    interf = report_data.get("interferences") or {}
    titulo = info.get("titulo_secao_9") or "RELATO DE INTERFERÊNCIAS EM CAMPO"
    blocks = [M.heading(titulo, level=1)]
    ocorrencias = interf.get("occurrences")
    if ocorrencias:
        for oc in ocorrencias:
            blocks.append(M.paragraph(str(oc), status="HITL",
                                      note="conteúdo real exige revisão humana (miolo descritivo)"))
    else:
        blocks.append(M.paragraph(
            _t("s9_negativa"), status="HITL",
            provenance=_prov("interferences"),
            note="ausência nas fontes consultadas não prova operacionalmente que não "
                "houve interferência (Q-10, FASE3_SLICE_3_SPEC.md §1) -- automation_class=TPL"))
    return M.Section(id="s9_interferencias", number="9", title=titulo, level=1,
                     page_setup="A4 retrato", blocks=blocks)


def _secao10_termo(report_data, info):
    """Implementa a logica inteira (contagem -> extenso -> frase), mas com
    n_paginas ainda nao comprovado (depende da montagem final -- Checkpoint
    3.3): fica PLACEHOLDER, nunca copia as paginas do golden."""
    period = report_data["period"]
    contract = report_data["contract"]
    road = report_data["road"]
    n_paginas = None  # provisorio -- so conhecido apos montar o documento final com anexos
    if n_paginas is None:
        n_txt, n_extenso_txt = "___", "___"
    else:
        n_txt = str(n_paginas)
        n_extenso_txt = por_extenso(n_paginas, feminino=True)
    texto = _fmt(
        "s10_termo", obra_sem_hifen=road.get("rodovia_sem_hifen") or info.get("rodovia_sem_hifen", "___"),
        # o numero do relatorio tem DUAS grafias no CURRENT, e nao e' inconsistencia
        # do golden: a capa usa 3 digitos com zeros a esquerda ("Relatorio: 049") e o
        # Termo usa o cardinal puro ("RD103 - N deg 49"). Conferido no RO_049.1.
        numero="%d" % report_data["report_number"],
        data_inicio=_data_br(period["start"]), data_fim=_data_br(period["end"]),
        contrato_obra=contract.get("contrato_obra") or "___",
        executora_contratual=road.get("empresa_executora_contrato") or "___",
        n_paginas=n_txt, n_paginas_extenso=n_extenso_txt,
        cidade=C.road_registry().get("meta", {}).get("cidade_emissao", "___"),
        data_emissao_extenso="___ (data de emissão só é conhecida na montagem final)")
    blocks = [
        M.heading("TERMO DE ENCERRAMENTO", level=1),
        M.paragraph(texto, status="PLACEHOLDER",
                   provenance=_prov("period") + _prov("contract.contrato_obra"),
                   note="n_paginas_provisorio=true -- Termo definitivo é gerado no "
                       "Checkpoint 3.3, após a montagem final (nunca copiado do golden)"),
    ]
    return M.Section(id="s10_termo_encerramento", number="10", title="TERMO DE ENCERRAMENTO",
                     level=1, page_setup="A4 retrato", blocks=blocks)


_TITULOS_ANEXO = {
    "ata_reuniao": "Ata de Reunião",
    "parecer_tecnico": "Parecer Técnico",
    "acompanhamento_ambiental": "Acompanhamento Ambiental",
    "acompanhamento_sst": "Acompanhamento de Saúde e Segurança do Trabalho",
    "nao_conformidade": "Não Conformidade",
    "diagrama_unifilar": "Diagrama Unifilar",
    "diario_de_obras": "Diário de Obras",
}
_ROMANOS = ["I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X"]


def _secao11_anexos(report_data):
    """Fase 3.4 (v0.2), CONFIG FIX VIS-07: o CURRENT usa parágrafos simples
    "Anexo N – Título", alinhados à margem, sem marcador de lista -- não a
    lista com bullet "•" do python-docx padrão. Cada item vira um Block
    `paragraph` com `role="anexos_item"`, que o renderer v2 trata sem estilo
    de lista e com respiro entre parágrafos (config/visual_style_v3.yaml
    `anexos.espaco_depois_pt`, ~28pt do GOLDEN_VISUAL_SPEC.md §8)."""
    annexes = report_data.get("annexes") or []
    blocks = [M.heading("ANEXOS", level=1)]
    for i, a in enumerate(annexes):
        titulo = _TITULOS_ANEXO.get(a["id"], a["id"])
        extra = ""
        if a["id"] == "parecer_tecnico" and a.get("pa_numero"):
            extra = " %s" % a["pa_numero"]
        romano = _ROMANOS[i] if i < len(_ROMANOS) else str(i + 1)
        item = M.paragraph("Anexo %s – %s%s" % (romano, titulo, extra),
                           provenance=_prov("annexes"))
        item.role = "anexos_item"
        blocks.append(item)
    return M.Section(id="s11_anexos", number="11", title="ANEXOS", level=1,
                     page_setup="A4 retrato", blocks=blocks)


# -----------------------------------------------------------------------------
# construcao completa
# -----------------------------------------------------------------------------
def resolver_modo(report_data, modo=None):
    """Modo de geração efetivo desta montagem.

    Ordem: argumento explícito > report_data.metadata.generation_mode >
    inferência pela distância entre o fim do período e hoje. O valor
    resolvido é gravado de volta no report_data, para que o manifesto e a
    comparação leiam o modo REALMENTE usado em vez de inferir de novo (e
    possivelmente diferente, num outro dia).
    """
    if modo is None:
        modo = (report_data.get("metadata") or {}).get("generation_mode")
    if modo is None:
        modo = C.inferir_modo((report_data.get("period") or {}).get("end"))
    report_data.setdefault("metadata", {})["generation_mode"] = modo
    return modo


def construir_documento(report_data, figuras=None, modo=None):
    """`figuras`: dict OPCIONAL {"ficha_acumulada": FiguraResult|None,
    "avanco_grafico": FiguraResult|None, "painel_fotografico":
    {servico_key: FiguraResult}, "ambiental": {"indices"|"condicionantes"|
    "qrcode": FiguraResult|None}, "nc": {"notificacoes"|"grafico_status"|
    "grafico_natureza": FiguraResult|None}} -- já resolvido pelo CHAMADOR
    (pipeline, ver shadow_report.cmd_candidate) via `figuras_bridge.tentar(...)`.
    content_builder.py continua sem conhecer caminho de arquivo/planilha
    (regra Sec.4 do docstring do módulo): quando `figuras` é None (default;
    é o que 'build'/'build2' e todos os testes existentes continuam
    passando), cada figura fica no placeholder de sempre -- comportamento
    IDÊNTICO ao pré-Fase 3.3."""
    figuras = figuras or {}
    modo = resolver_modo(report_data, modo)
    obra = report_data["metadata"]["road_key"]
    info = C.obra_info(obra)
    versoes = C.report_versions()
    v3 = versoes["report_versions"]["v3"]
    titulo_4 = info.get("titulo_secao_4") or "DIAGRAMA UNIFILAR"

    namer = ServiceNamer()  # 1a ocorrencia compartilhada entre s1 e s4 (mesmo documento)
    seq = CaptionSequencer()  # Sec.3.1: sequencias Figura/Gráfico/Tabela do documento inteiro

    sections = [
        _secao_capa(report_data, info),
        _secao_indice(v3["secoes"]),
        _secao1_apresentacao(report_data, info, namer),
        _secao2_ficha(report_data, seq, figuras.get("ficha_acumulada")),
        _secao3_avanco(report_data, seq, figuras.get("avanco_grafico"), info),
        _secao4_unifilar(report_data, info, namer, titulo_4),
    ]
    sections.extend(_secao5_execucao(report_data, info, seq, figuras.get("painel_fotografico")))
    sections.append(_secao6_ambiental(report_data, info, seq, figuras.get("ambiental"), modo))
    sections.append(_secao7_diario(report_data, info))
    sections.append(_secao8_nc(report_data, info, seq, figuras.get("nc")))
    sections.append(_secao9_interferencias(report_data, info))
    sections.append(_secao10_termo(report_data, info))
    sections.append(_secao11_anexos(report_data))

    meta = {
        "road_key": obra, "road": report_data["road"]["rodovia"],
        "report_number": report_data["report_number"],
        "period_start": report_data["period"]["start"],
        "period_end": report_data["period"]["end"],
        "schema_version": report_data["metadata"]["schema_version"],
        "generated_at": dt.datetime.now().astimezone().isoformat(timespec="seconds"),
        "generator": "scripts/gerador/content_builder.py (Fase 3, Slice 3)",
        "aviso": "RELATORIO-SOMBRA -- gerado em paralelo para validacao; nao substitui o oficial",
    }
    return M.Document(meta=meta, sections=sections)
