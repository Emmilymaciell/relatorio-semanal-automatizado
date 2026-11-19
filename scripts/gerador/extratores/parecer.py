# -*- coding: utf-8 -*-
r"""
scripts/gerador/extratores/parecer.py - Q-11 (spec Sec.1): associa o Parecer
Tecnico (PA-NNN) a semana do relatorio LENDO O TEXTO do PDF, nunca por mtime,
maior numero ou "PA anterior + 1" (a numeracao PA-NNN e corporativa, nao por
obra -- ver evidencia em report_schema.yaml s11_anexos.parecer_tecnico).

Pipeline (verificado diretamente para RD103/2026-08-29..09-04, sessao
2026-09-17, contra os 11 PDFs de
'6. Acompanhamento ambiental e SST\Parecer\' da obra):

    candidatos: todo arquivo 'PA-NNN[ -]Parecer.pdf' da pasta Parecer da obra
        -> ler o TEXTO do PDF (PyMuPDF/fitz)
        -> extrair rodovia mencionada + data(s) de verificacao em campo
        -> casar rodovia == obra E data de campo dentro do periodo
    0 compativeis -> NOT_FOUND
    1 compativel  -> MATCH
    >1 compativeis -> HITL (lista todos)

Sinal de rodovia: qualquer trecho do texto que bata com 'RD[- ]?NNN[ ]?A?',
normalizado (sem espaco/hifen, maiusculo) e comparado a
road_registry.obras.<obra>.rodovia_sem_hifen.

Sinal de periodo: a data de verificacao em campo e a que aparece PROXIMA da
palavra 'campo' no texto (ex.: "verificacoes realizadas em campo em 2 de
setembro de 2026"), nao qualquer data solta do documento (o cabecalho tem uma
data de EMISSAO do parecer, que fica na semana SEGUINTE a vistoria -- nao
serve para casar com o periodo do relatorio). Verificado nos 11 candidatos da
RD103: dos 11, so o PA-254 tem uma data-de-campo dentro de 2026-08-29..09-04;
todos os outros (PA-096/125/189/201/202/203/205/217/220/259) ou nao tem essa
frase, ou a data fica fora do periodo -- o mecanismo chega em PA-254 sozinho,
sem qualquer atalho por numero/mtime.
"""
import os
import re
import sys
import datetime as dt

AQUI = os.path.dirname(os.path.abspath(__file__))
SCRIPTS = os.path.dirname(os.path.dirname(AQUI))
if SCRIPTS not in sys.path:
    sys.path.insert(0, SCRIPTS)
import prog_util as U  # noqa: E402
from .. import assert_nao_e_golden  # noqa: E402
from .. import config as C  # noqa: E402

_RE_ARQUIVO_PARECER = re.compile(r"^PA-(\d{3})\s*-?\s*Parecer\.pdf$", re.IGNORECASE)

_MESES_PT = {
    "janeiro": 1, "fevereiro": 2, "marco": 3, "março": 3, "abril": 4, "maio": 5,
    "junho": 6, "julho": 7, "agosto": 8, "setembro": 9, "outubro": 10,
    "novembro": 11, "dezembro": 12,
}
_RE_DATA_EXTENSO = re.compile(
    r"(\d{1,2})\s*[ºo°]?\s+de\s+(" + "|".join(_MESES_PT) + r")\s+de\s+(\d{4})", re.IGNORECASE)
# data-de-campo: a data por extenso tem que aparecer "perto" (<=60 caracteres)
# da palavra 'campo' -- e o que distingue "verificacao(oes) ... em campo ...
# em <data>" (data da VISTORIA) de qualquer outra data solta do documento
# (ex.: data de emissao da licenca, data do proximo relatorio, etc.)
_RE_DATA_DE_CAMPO = re.compile(r"campo[^.]{0,60}?" + _RE_DATA_EXTENSO.pattern, re.IGNORECASE)

_RE_RODOVIA_NO_TEXTO = re.compile(r"RD[\s-]?\d{2,3}\s?A?", re.IGNORECASE)


def _normalizar_rodovia(s):
    return re.sub(r"[^A-Za-z0-9]", "", s).upper()


def localizar_pasta_parecer(obra):
    """{obra}\\3_Gestao Tecnica de Obra\\6. Acompanhamento ambiental e SST\\Parecer.
    Resolvido por prefixo/igualdade normalizados (config/source_layout.yaml
    'pasta_ambiental_sst' + 'subpastas_fixas: [Parecer]'), nunca por caminho
    literal por obra."""
    pasta_gto = os.path.join(U.BASE_OBRAS, obra, "3_Gestao Tecnica de Obra")
    pasta_amb_sst = C.resolver_subpasta(pasta_gto, prefixo="6. acompanhamento ambiental")
    if not pasta_amb_sst:
        return None
    return C.resolver_subpasta(pasta_amb_sst, iguais=["parecer"])


def listar_candidatos(obra):
    """Todo 'PA-NNN[ -]Parecer.pdf' da pasta Parecer da obra -- SO metadados
    de arquivo (nome/numero/caminho), sem ler texto ainda."""
    pasta = localizar_pasta_parecer(obra)
    if not pasta or not U.exists(pasta):
        return [], pasta
    candidatos = []
    for nome in U.listdir(pasta):
        caminho = os.path.join(pasta, nome)
        if U.isdir(caminho):
            continue
        m = _RE_ARQUIVO_PARECER.match(nome)
        if not m:
            continue
        candidatos.append({"numero": "PA-%s" % m.group(1), "nome": nome, "path": caminho})
    candidatos.sort(key=lambda c: c["numero"])
    return candidatos, pasta


def extrair_texto_pdf(caminho):
    """Copia para tmp/ (nunca abre o original) e le o texto de TODAS as
    paginas com PyMuPDF. Chama a guarda anti-golden antes de abrir -- os
    PA-NNN nunca batem com o padrao do golden, mas todo extrator tem que
    chamar a guarda (spec Sec.5/0)."""
    assert_nao_e_golden(caminho)
    if not U.exists(caminho):
        return None
    import fitz  # import local: só quem lê PDF de verdade paga o custo
    destino = "parecer_%s.pdf" % U.slug(os.path.basename(caminho))
    local = U.copiar_para_tmp(caminho, "parecer", destino)
    doc = fitz.open(local)
    try:
        return "".join(pagina.get_text() for pagina in doc)
    finally:
        doc.close()


def extrair_sinais(texto):
    """rodovias mencionadas (normalizadas) + datas-de-campo (com trecho de
    evidencia) + paragrafo 'DO OBJETO' (so para citação/evidencia, nao entra
    na decisao)."""
    rodovias = {_normalizar_rodovia(m.group(0)) for m in _RE_RODOVIA_NO_TEXTO.finditer(texto)}

    datas_campo = []
    for m in _RE_DATA_DE_CAMPO.finditer(texto):
        dia, mes_nome, ano = m.group(1), m.group(2), m.group(3)
        mes = _MESES_PT.get(mes_nome.lower())
        try:
            data = dt.date(int(ano), mes, int(dia))
        except ValueError:
            continue
        trecho = re.sub(r"\s+", " ", m.group(0)).strip()
        datas_campo.append({"data": data.isoformat(), "trecho": trecho})

    objeto = None
    m = re.search(r"DO OBJETO\s*\n?(.*?)\n\d+\.\s", texto, re.DOTALL)
    if m:
        objeto = re.sub(r"\s+", " ", m.group(1)).strip()[:800]

    return {"rodovias_normalizadas": rodovias, "datas_campo": datas_campo, "objeto": objeto}


def avaliar(obra, data_inicio, data_fim):
    """Executa o pipeline completo (spec Sec.1/Q-11) e devolve:
        {"status": "NOT_FOUND"|"MATCH"|"HITL",
         "candidatos_analisados": [...],           # todo PA da pasta, com os sinais lidos
         "documento": {...} | None,                # so quando status == MATCH
         "candidatos": [...] | None,                # so quando status == HITL
         "mecanismo": "...",                        # descricao do metodo usado
         "motivo": "..."}                           # so quando NOT_FOUND/HITL
    """
    if isinstance(data_inicio, str):
        data_inicio = dt.date.fromisoformat(data_inicio)
    if isinstance(data_fim, str):
        data_fim = dt.date.fromisoformat(data_fim)

    info = C.obra_info(obra)
    rodovia_alvo = _normalizar_rodovia(info["rodovia_sem_hifen"])

    candidatos, pasta = listar_candidatos(obra)
    mecanismo = ("leitura do TEXTO de cada PDF PA-NNN da pasta Parecer da obra: casa rodovia "
                 "mencionada no texto (normalizada) com a rodovia da obra E data de verificacao "
                 "em campo (padrao 'campo ... em <data por extenso>') dentro do periodo do "
                 "relatorio -- nunca mtime, maior numero ou PA anterior+1")

    if not pasta:
        return {"status": "NOT_FOUND", "candidatos_analisados": [], "documento": None,
                "candidatos": None, "mecanismo": mecanismo,
                "motivo": "pasta 'Parecer' nao encontrada em 6. Acompanhamento ambiental e SST "
                          "para %s" % obra}
    if not candidatos:
        return {"status": "NOT_FOUND", "candidatos_analisados": [], "documento": None,
                "candidatos": None, "mecanismo": mecanismo,
                "motivo": "pasta 'Parecer' existe mas nenhum arquivo 'PA-NNN - Parecer.pdf' "
                          "foi encontrado nela (%s)" % pasta}

    analisados = []
    compativeis = []
    for c in candidatos:
        texto = extrair_texto_pdf(c["path"])
        if texto is None:
            analisados.append(dict(c, status="UNREADABLE"))
            continue
        sinais = extrair_sinais(texto)
        rodovia_ok = rodovia_alvo in sinais["rodovias_normalizadas"]
        datas_no_periodo = [d for d in sinais["datas_campo"]
                            if data_inicio <= dt.date.fromisoformat(d["data"]) <= data_fim]
        periodo_ok = bool(datas_no_periodo)
        item = dict(c, status="OK", rodovia_ok=rodovia_ok, periodo_ok=periodo_ok,
                   datas_campo_no_periodo=datas_no_periodo,
                   datas_campo_no_texto=sinais["datas_campo"], objeto=sinais["objeto"])
        analisados.append(item)
        if rodovia_ok and periodo_ok:
            compativeis.append(item)

    if len(compativeis) == 0:
        return {"status": "NOT_FOUND", "candidatos_analisados": analisados, "documento": None,
                "candidatos": None, "mecanismo": mecanismo,
                "motivo": "nenhum PA da pasta bate simultaneamente rodovia (%s) + data de "
                          "verificacao em campo dentro de %s..%s" % (
                              info["rodovia"], data_inicio.isoformat(), data_fim.isoformat())}

    if len(compativeis) == 1:
        doc = compativeis[0]
        evidencia = "%s (%s): %s" % (
            doc["numero"], os.path.basename(doc["path"]),
            "; ".join("'%s' -> %s" % (d["trecho"], d["data"]) for d in doc["datas_campo_no_periodo"]))
        return {"status": "MATCH", "candidatos_analisados": analisados, "documento": doc,
                "candidatos": None, "mecanismo": mecanismo, "evidencia": evidencia}

    return {"status": "HITL", "candidatos_analisados": analisados, "documento": None,
            "candidatos": compativeis, "mecanismo": mecanismo,
            "motivo": "mais de um PA compativel (rodovia + periodo) -- confirmacao humana "
                      "necessaria: %s" % ", ".join(c["numero"] for c in compativeis)}
