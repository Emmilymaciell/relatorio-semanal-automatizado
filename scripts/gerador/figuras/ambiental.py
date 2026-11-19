# -*- coding: utf-8 -*-
r"""
scripts/gerador/figuras/ambiental.py - Paineis ambientais (Figura 6 "indices",
Figura 7 "condicionantes") + QR Code (Figura 8).

Fonte: `04_Acompanhamento - Licenca Ambiental.xlsx`, abas `Empresas`,
`Licencas` e `Condicionantes` (spec Sec.2.4). O caminho vem do manifest
(`sources.environmental.licenca_ambiental_planilha.path`), nunca hardcoded.

## SHADOW_ERROR corrigido nesta versao (v0.2): filtro por RODOVIA, nao empresa

A v0.1 filtrava `Licencas`/`Condicionantes` pela EMPRESA EXECUTORA do
CONTRATO (`road.empresa_executora_contrato`, ex.: "Consórcio Executora B"
para a RD-102A). Duas falhas reais, medidas:

1. A licenca ambiental e' emitida para a RODOVIA, nao para a empresa -- uma
   obra pode trocar de executora ao longo do contrato sem que a Licenca
   Ambiental Unica mude de titular. Filtrar por empresa e' a premissa
   errada (registrado em `docs/fontes/dashboard_ambiental_powerbi.md` --
   o dashboard Power BI filtra por Rodovia, nao por empresa executora).
2. Na pratica isso ja QUEBRAVA a RD-102A: a razao social da planilha
   (`EXECUTORA B`) nao tem nenhuma substring em
   comum com o nome do CONTRATO (`Consórcio Executora B`) -- o filtro
   antigo devolvia sempre `LIMITATION: nao ha linha de ... na aba
   'Licencas'` para essa obra, silenciosamente.

O UNICO elo obra<->dado ambiental na planilha e' a aba `Empresas` (coluna
"Endereco", que na pratica guarda a Rodovia -- ver
`docs/fontes/acompanhamento_ambiental.md` Sec.1.1). Este modulo le essa
aba primeiro e resolve "quais razoes sociais atendem esta rodovia HOJE",
depois filtra `Licencas`/`Condicionantes` por essas razoes sociais -- nunca
por comparacao de string com o nome do contrato, e nunca com `if obra==`
(a chave e' sempre a rodovia recebida por parametro, resolvida contra a
planilha, igual para as 4 obras).

## Os dois niveis do dashboard (Tarefa 1, `dashboard_ambiental_powerbi.md`)

O painel Power BI publicado mistura dois recortes:
  - `corporate_kpis`: os totais das 4 rodovias somadas (o que aparece
    quando ninguem clicou em nenhum filtro de rodovia);
  - `road_specific_metrics`: cartoes/barra/tabela filtrados pela rodovia
    de UMA obra (o que aparece quando o slicer de Rodovia esta ativo).

Comparando os 8 relatorios oficiais (4 obras x 2 semanas) na MESMA semana,
ficou `CONFIRMADO POR EVIDENCIA DIRETA` (ver a secao "Filtro - resolvido
documentalmente" no arquivo acima) que o screenshot que vai para a secao 6
de CADA relatorio e' sempre o **road_specific_metrics da propria obra**
(slicer de rodovia realcado, barra unica, cartoes diferentes por obra) --
NUNCA o `corporate_kpis` sem filtro. Por isso as duas funcoes abaixo:
  - calculam e devolvem os DOIS niveis (`corporate_kpis` e
    `road_specific_metrics`, ambos em `provenance`) -- "separar os dois
    niveis no modelo", como pede a Tarefa 2;
  - mas DESENHAM o painel (o PNG) a partir de `road_specific_metrics`,
    porque e' esse o recorte que o relatorio real usa.

## SOURCE_LIMITATION (mesma familia do TODAY() das NC)

A copia local da planilha (`docProps/modified` de 2026-05-13, ver
`acompanhamento_ambiental.md` Sec.2.5) fica desatualizada frente ao Power
BI, que le a planilha corrente (atualizada continuamente). Registros novos
simplesmente NAO EXISTEM na copia local -- nenhuma logica de filtro resolve
isso. Os numeros aqui sao um RECALCULO REAL e VERIFICAVEL sobre os dados
que a copia local tem agora, nao uma tentativa de reproduzir o estado de
uma semana passada (isso seria golden como fonte de dado, proibido) nem de
bater com o numero mais recente visto no painel publicado (idem). Igual ao
`TODAY()` das NC: o estado e' sempre o ATUAL da fonte disponivel, marcado
como tal.

## Status das condicionantes: taxonomia com 3+ valores (nao 2)

`Condicionantes!G` e' uma formula que usa `HOJE()` (documentada em
`acompanhamento_ambiental.md` Sec.1.3): `faixa(x) = "Pendente" se x<=0;
"Proximo do vencimento" se x<=30; senao "No prazo"`, mais "Finalizada"
quando a coluna Observacoes diz isso. O valor CACHEADO na copia local foi
calculado pelo Excel com o `HOJE()` de 2026-05-13 e por isso, para varias
empresas, no' momento vem so' com 2 categorias presentes (sorte de
distribuicao de datas, nao decisao de codigo). Este modulo REIMPLEMENTA a
formula (`_status_condicionante`) usando a data REAL de hoje (a mesma
filosofia do `nc_graficos.py`/`extratores/nc.py` para `TODAY()`), o que
recupera as 3-4 categorias reais (`Finalizada`, `No prazo`, `Proximo do
vencimento`, e `Pendente` quando houver) em vez de depender de um cache
frio. Nao e' "forcar bater com o golden": e' a MESMA formula documentada,
so' que calculada com a data de execucao real, nunca com uma data
escolhida a dedo para produzir um numero especifico.

## QR Code

Bibliotecas `matplotlib` e `qrcode` nao existem neste ambiente (spec Sec.0.1)
e nao podem ser instaladas. Emitimos um placeholder quadrado 79,07x79,07mm
com `status=PRESENTATION_DEPENDENCY`, incluido no relatorio e claramente
marcado como tal -- nunca um QR Code falso apresentado como real.
"""
import os
import datetime as dt
from collections import Counter

import openpyxl

from . import (FiguraResult, mm_to_px, get_font, wrap_text, novo_canvas,
               salvar_png, desenhar_borda, text_width, COR_BORDA, COR_PRETO,
               COR_SHADING_CABECALHO, U, C, assert_nao_e_golden)

LARGURA_MM = 160.0
QR_LADO_MM = 79.07

# Ordem de exibicao da taxonomia de condicionantes (Sec.1.3 do doc-fonte).
# So' um vocabulario fixo para ordenar a tabela -- nao filtra nem descarta
# nenhuma categoria que aparecer nos dados (se surgir uma 5a categoria nova
# na planilha, ela ainda aparece na tabela, so' que ao final).
_ORDEM_STATUS_CONDICIONANTE = ["Pendente", "Próximo do vencimento", "No prazo", "Finalizada"]


def _norm_empresa(s):
    if not s:
        return ""
    n = U.norm(s)
    for sufixo in (" ltda", " s/a", " s.a.", " - ccl", " - ccb"):
        n = n.replace(sufixo, "")
    return n.strip()


def _norm_rodovia(s):
    """Normaliza rotulo de rodovia para comparacao: sem acento/maiuscula (via
    U.norm) e sem espacos/hifens -- 'RD-102 A' (road_registry) e 'RD-102a'
    (planilha, aba Empresas) tem que colidir no mesmo valor."""
    if not s:
        return ""
    n = U.norm(s)
    return n.replace("-", "").replace(" ", "")


def _abrir(caminho):
    assert_nao_e_golden(caminho)
    destino = "ambiental_%s.xlsx" % U.slug(os.path.basename(caminho))
    local = U.copiar_para_tmp(caminho, "figuras_ambiental", destino)
    return openpyxl.load_workbook(local, data_only=True)


def _ler_empresas_por_rodovia(wb):
    """Le a aba 'Empresas': unico elo obra<->dado ambiental na planilha
    (coluna 'Endereco', que na pratica guarda a Rodovia -- confirmado em
    `acompanhamento_ambiental.md` Sec.1.1). Devolve
    (rodovia_norm -> [razao_social,...], erro)."""
    nomes = [n for n in wb.sheetnames if U.norm(n).startswith("empresa")]
    if not nomes:
        return None, "aba 'Empresas' nao encontrada (abas: %s)" % wb.sheetnames
    ws = wb[nomes[0]]
    header_row = _achar_header_row(ws)
    headers = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(header_row, c).value
        if isinstance(v, str) and v.strip():
            headers[U.norm(v)] = c
    c_razao = headers.get("razao social")
    # a coluna se chama "Endereço" mas guarda a rodovia (achado documentado,
    # nao inferencia deste modulo).
    c_rodovia = headers.get("endereco")
    if not c_razao or not c_rodovia:
        return None, ("colunas 'Razão social'/'Endereço' nao encontradas no cabecalho "
                      "da aba 'Empresas' (linha %d)" % header_row)
    por_rodovia = {}
    for r in range(header_row + 1, ws.max_row + 1):
        razao = ws.cell(r, c_razao).value
        rodovia = ws.cell(r, c_rodovia).value
        if razao and rodovia:
            por_rodovia.setdefault(_norm_rodovia(rodovia), []).append(str(razao).strip())
    return por_rodovia, (None if por_rodovia else "aba 'Empresas' sem nenhuma linha valida")


def _achar_header_row(ws, max_r=5, max_c=20):
    """Acha a 1a linha com >=3 celulas de texto nao vazias (header real —
    a linha 1 costuma ser so um titulo mesclado)."""
    for r in range(1, max_r + 1):
        n_texto = sum(1 for c in range(1, max_c + 1)
                     if isinstance(ws.cell(r, c).value, str) and ws.cell(r, c).value.strip())
        if n_texto >= 3:
            return r
    return 1


def _ler_licencas_todas(wb):
    """Le TODAS as linhas da aba 'Licencas' (sem filtro) -- o filtro por
    rodovia e' aplicado depois, em `_agrupar_licencas_por_rodovia`, contra a
    razao social resolvida pela aba 'Empresas'."""
    nomes = [n for n in wb.sheetnames if U.norm(n).startswith("licen")]
    if not nomes:
        return None, "aba 'Licenças' nao encontrada (abas: %s)" % wb.sheetnames
    ws = wb[nomes[0]]
    header_row = _achar_header_row(ws)
    headers = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(header_row, c).value
        if isinstance(v, str) and v.strip():
            headers[U.norm(v)] = c
    c_emp = headers.get("empresa")
    if not c_emp:
        return None, "coluna 'Empresa' nao encontrada no cabecalho (linha %d)" % header_row
    # "Nº Licença" normaliza para "no licenca" (U.norm decompoe o indicador
    # ordinal 'º' em 'o' -- nao ha combining char para remover, entao o 'o'
    # fica). Aceita as duas formas para nao quebrar se a grafia mudar.
    c_num = headers.get("no licenca") or headers.get("n licenca") or c_emp
    linhas = []
    for r in range(header_row + 1, ws.max_row + 1):
        emp = ws.cell(r, c_emp).value
        if not emp:
            continue
        linhas.append({
            "empresa": str(emp).strip(),
            "n_licenca": ws.cell(r, c_num).value,
            "fase": ws.cell(r, headers.get("fase da licenca", c_emp)).value,
            "emissao": ws.cell(r, headers.get("data de emissao", c_emp)).value,
            "validade": ws.cell(r, headers.get("data de validade", c_emp)).value,
            "orgao": ws.cell(r, headers.get("orgao", c_emp)).value,
            "status": ws.cell(r, headers.get("status", c_emp)).value,
            "vence_em": ws.cell(r, headers.get("vence em", c_emp)).value,
        })
    return linhas, (None if linhas else "aba 'Licenças' sem nenhuma linha valida")


def _ler_condicionantes_todas(wb):
    """Le TODAS as linhas da aba 'Condicionantes' (sem filtro), com os campos
    crus necessarios para reimplementar `_status_condicionante` (data de
    entrega + observacoes) em vez de confiar so' no valor cacheado."""
    nomes = [n for n in wb.sheetnames if U.norm(n).startswith("condicionante")]
    if not nomes:
        return None, "aba 'Condicionantes' nao encontrada (abas: %s)" % wb.sheetnames
    ws = wb[nomes[0]]
    header_row = _achar_header_row(ws)
    headers = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(header_row, c).value
        if isinstance(v, str) and v.strip():
            headers[U.norm(v)] = c
    c_resp = headers.get("responsavel")
    c_status_cache = headers.get("status")
    c_entrega = headers.get("data da entrega") or headers.get("data de entrega")
    c_obs = headers.get("observacoes") or headers.get("observações")
    if not c_resp or not c_status_cache:
        return None, ("colunas 'Responsável'/'Status' nao encontradas no cabecalho "
                      "(linha %d)" % header_row)
    linhas = []
    for r in range(header_row + 1, ws.max_row + 1):
        resp = ws.cell(r, c_resp).value
        if not resp:
            continue
        linhas.append({
            "responsavel": str(resp).strip(),
            "status_cache": ws.cell(r, c_status_cache).value,
            "data_entrega": ws.cell(r, c_entrega).value if c_entrega else None,
            "observacoes": ws.cell(r, c_obs).value if c_obs else None,
        })
    return linhas, (None if linhas else "aba 'Condicionantes' sem nenhuma linha valida")


def _empresas_da_rodovia(por_rodovia, rodovia_alvo):
    """`por_rodovia`: dict rodovia_norm -> [razao_social,...] (de
    `_ler_empresas_por_rodovia`). Devolve o conjunto de razoes sociais
    (normalizadas para comparacao) que atendem `rodovia_alvo` HOJE."""
    return set(_norm_empresa(e) for e in por_rodovia.get(_norm_rodovia(rodovia_alvo), []))


def _status_condicionante(data_entrega, observacoes, hoje):
    """Reimplementa `Condicionantes!G` (formula compartilhada, documentada em
    `docs/fontes/acompanhamento_ambiental.md` Sec.1.3), com `HOJE()`
    substituido pelo parametro `hoje` (a data real de execucao, nunca uma
    data escolhida para bater com outro numero):

        J = SE(H="Atendida";"Atendida"; SE(H="Finalizada";"Finalizada"; F-HOJE()))
        K = SE(J="Atendida"; DATA(ANO(F)+1;MES(F);DIA(F)); SE(H="Finalizada";...;F))
        L = SE(J="Atendida"; K-HOJE(); J)
        G = SE(J="Finalizada";"Finalizada"; SE(J="Atendida"; faixa(L); faixa(J)))
        faixa(x) = "Pendente" se x<=0; "Próximo do vencimento" se x<=30; senao "No prazo"

    Devolve None quando nao ha 'Data da entrega' para calcular (linha
    incompleta) -- nesse caso o chamador cai de volta no valor cacheado da
    planilha (`status_cache`), nunca inventa um status.
    """
    obs = str(observacoes).strip() if observacoes else ""
    if obs == "Finalizada":
        return "Finalizada"
    if data_entrega is None:
        return None
    entrega = data_entrega.date() if hasattr(data_entrega, "date") else data_entrega
    if not hasattr(entrega, "toordinal"):
        return None
    if obs == "Atendida":
        try:
            alvo = entrega.replace(year=entrega.year + 1)
        except ValueError:
            # 29/fev sem ano bissexto no ano seguinte
            alvo = entrega.replace(year=entrega.year + 1, day=entrega.day - 1)
        dias = (alvo - hoje).days
    else:
        dias = (entrega - hoje).days
    if dias <= 0:
        return "Pendente"
    if dias <= 30:
        return "Próximo do vencimento"
    return "No prazo"


def _kpis_licencas(linhas, por_rodovia):
    """Separa os DOIS niveis do modelo (Tarefa 2, item 2) a partir de TODAS
    as linhas de 'Licencas' + o mapa `por_rodovia` da aba 'Empresas':

      - corporate_kpis: totais somando as 4 rodovias (o que o Power BI
        mostra sem nenhum filtro de rodovia ativo);
      - road_specific_metrics: por rodovia (o que o Power BI mostra com o
        slicer de rodovia ativo -- e' o recorte que o relatorio real usa,
        ver docstring do modulo)."""
    empresa_para_rodovia = {}
    for rod_norm, empresas in por_rodovia.items():
        for e in empresas:
            empresa_para_rodovia[_norm_empresa(e)] = rod_norm

    por_rodovia_norm = {}
    nao_resolvidas = Counter()
    for l in linhas:
        rod = empresa_para_rodovia.get(_norm_empresa(l["empresa"]))
        if rod is None:
            nao_resolvidas[l["empresa"]] += 1
            continue
        por_rodovia_norm.setdefault(rod, []).append(l)

    def _resumo(ls):
        return {
            "total_licencas": len(ls),
            "emitidas": sum(1 for l in ls if (l["status"] or "").strip() == "Vigente"),
            "pendentes": sum(1 for l in ls
                             if (l["status"] or "").strip() == "Aguardando a liberação"),
            "fases": dict(Counter((l["fase"] or "(sem fase)") for l in ls)),
        }

    corporate = _resumo(linhas)
    corporate["total_licencas_por_rodovia"] = {
        rod: len(ls) for rod, ls in por_rodovia_norm.items()}
    corporate["empresas_nao_resolvidas_para_rodovia"] = dict(nao_resolvidas)

    road_specific = {rod: dict(_resumo(ls), linhas=ls) for rod, ls in por_rodovia_norm.items()}
    return corporate, road_specific


def _kpis_condicionantes(linhas, por_rodovia, hoje):
    """Mesma separacao de `_kpis_licencas`, para 'Condicionantes'. O status
    de cada linha e' recalculado com `_status_condicionante` (taxonomia
    completa, ver docstring do modulo); quando a linha nao tem dado
    suficiente para recalcular, cai no `status_cache` da planilha."""
    empresa_para_rodovia = {}
    for rod_norm, empresas in por_rodovia.items():
        for e in empresas:
            empresa_para_rodovia[_norm_empresa(e)] = rod_norm

    linhas_calc = []
    for l in linhas:
        status = _status_condicionante(l["data_entrega"], l["observacoes"], hoje)
        if status is None:
            status = str(l["status_cache"] or "(vazio)").strip()
        linhas_calc.append(dict(l, status_recalculado=status))

    por_rodovia_norm = {}
    nao_resolvidas = Counter()
    for l in linhas_calc:
        rod = empresa_para_rodovia.get(_norm_empresa(l["responsavel"]))
        if rod is None:
            nao_resolvidas[l["responsavel"]] += 1
            continue
        por_rodovia_norm.setdefault(rod, []).append(l)

    def _resumo(ls):
        total = len(ls)
        contagem = Counter(l["status_recalculado"] for l in ls)
        return {"total_condicionantes": total, "contagem_por_status": dict(contagem)}

    corporate = _resumo(linhas_calc)
    corporate["total_condicionantes_por_rodovia"] = {
        rod: len(ls) for rod, ls in por_rodovia_norm.items()}
    corporate["responsaveis_nao_resolvidos_para_rodovia"] = dict(nao_resolvidas)

    road_specific = {rod: dict(_resumo(ls), linhas=ls) for rod, ls in por_rodovia_norm.items()}
    return corporate, road_specific


def _fmt_data(v):
    try:
        return v.strftime("%d/%m/%Y")
    except AttributeError:
        return str(v) if v is not None else ""


def _fmt_dias(v):
    if isinstance(v, (int, float)):
        return "%d dias" % v
    if hasattr(v, "year"):
        # a planilha calcula "Vence em" como (validade - hoje); quando
        # 'Data de validade' esta vazia, o resultado e uma data-sentinela
        # sem sentido de negocio (ex.: ano 1773) -- nao e um numero de dias
        # nem uma data real de vencimento. Nao formatamos como se fosse:
        # mostramos "-" e preservamos o valor cru em nenhum outro lugar
        # (nao inventamos um numero de dias que a planilha nao calculou).
        if v.year < 2000:
            return "-"
        return v.strftime("%d/%m/%Y")
    return str(v) if v not in (None, "") else "-"


def _desenhar_tabela(titulo, colunas, linhas, largura_mm, altura_mm):
    """Tabela generica com cabecalho sombreado (estilo secao 5 do
    GOLDEN_VISUAL_SPEC.md), usada pelas Figuras 6 e 7."""
    img, draw = novo_canvas(largura_mm, altura_mm)
    w_px, h_px = img.size
    fonte_titulo = get_font(9.5, bold=True)
    fonte_cab = get_font(7.5, bold=True)
    fonte = get_font(7.5)

    titulo_h = mm_to_px(5.5)
    draw.text((mm_to_px(1.5), mm_to_px(0.8)), titulo, font=fonte_titulo, fill=COR_PRETO)

    tabela_y0 = titulo_h
    cabecalho_h = mm_to_px(5.5)
    linha_h = (h_px - tabela_y0 - cabecalho_h) / max(1, len(linhas))
    larguras_px = [mm_to_px(largura_mm * peso) for _, peso in colunas]
    larguras_px[-1] += w_px - sum(larguras_px)

    draw.rectangle([0, tabela_y0, w_px, tabela_y0 + cabecalho_h], fill=COR_SHADING_CABECALHO())
    x = 0
    for (nome, _peso), w in zip(colunas, larguras_px):
        tw = text_width(draw, nome, fonte_cab)
        draw.text((x + w / 2 - tw / 2, tabela_y0 + cabecalho_h / 2 - mm_to_px(1.3)),
                  nome, font=fonte_cab, fill=COR_PRETO)
        x += w

    linha_fonte_h = mm_to_px(3.2)
    max_linhas_celula = max(1, int((linha_h - mm_to_px(1.2)) // linha_fonte_h))
    y = tabela_y0 + cabecalho_h
    for linha in linhas:
        x = 0
        for valor, w in zip(linha, larguras_px):
            ls = wrap_text(draw, str(valor), fonte, w - mm_to_px(2))
            ty = y + mm_to_px(0.6)
            for lt in ls[:max_linhas_celula]:
                draw.text((x + mm_to_px(1), ty), lt, font=fonte, fill=COR_PRETO)
                ty += linha_fonte_h
            x += w
        y += linha_h

    desenhar_borda(draw, (0, tabela_y0, w_px - 1, h_px - 1), 1, COR_BORDA)
    x = 0
    for w in larguras_px[:-1]:
        x += w
        draw.line([(x, tabela_y0), (x, h_px)], fill=COR_BORDA, width=1)
    y = tabela_y0 + cabecalho_h
    draw.line([(0, y), (w_px, y)], fill=COR_BORDA, width=1)
    for _ in linhas:
        y += linha_h
        draw.line([(0, y), (w_px, y)], fill=COR_BORDA, width=1)
    return img


def gerar_indices(caminho_licenca_ambiental, rodovia, out_dir) -> FiguraResult:
    largura_mm, altura_mm = 160.0, 86.89  # medida do GOLDEN_VISUAL_SPEC.md Sec.6 (Figura 6)
    prov = {"source_file": caminho_licenca_ambiental, "sheet": "Licenças", "rodovia": rodovia}
    if not caminho_licenca_ambiental or not U.exists(caminho_licenca_ambiental):
        return FiguraResult(None, largura_mm, altura_mm, "LIMITATION",
                            "arquivo de licenca ambiental nao encontrado", prov)
    if not rodovia:
        return FiguraResult(None, largura_mm, altura_mm, "LIMITATION",
                            "rodovia nao informada (esperado road.rodovia do report_data)", prov)
    wb = _abrir(caminho_licenca_ambiental)
    por_rodovia_empresas, erro = _ler_empresas_por_rodovia(wb)
    if por_rodovia_empresas is None:
        return FiguraResult(None, largura_mm, altura_mm, "LIMITATION", erro, prov)
    linhas, erro = _ler_licencas_todas(wb)
    if linhas is None:
        return FiguraResult(None, largura_mm, altura_mm, "LIMITATION", erro, prov)

    corporate_kpis, road_specific_metrics = _kpis_licencas(linhas, por_rodovia_empresas)
    metrica_obra = road_specific_metrics.get(_norm_rodovia(rodovia))
    if not metrica_obra or not metrica_obra["linhas"]:
        return FiguraResult(None, largura_mm, altura_mm, "LIMITATION",
                            "nenhuma licenca resolvida para a rodovia '%s' (aba 'Empresas' "
                            "nao liga nenhuma razao social a essa rodovia, ou a rodovia "
                            "nao tem linha em 'Licenças')" % rodovia, prov)
    linhas_obra = metrica_obra["linhas"]

    # Tarefa 1 (dashboard_ambiental_powerbi.md, "Filtro - resolvido
    # documentalmente"): o screenshot que vai para a secao 6 e' sempre
    # filtrado pela rodovia da propria obra -- e' o recorte que se desenha
    # aqui (road_specific_metrics). corporate_kpis fica so' em provenance,
    # para rastreabilidade/uso futuro (Tarefa 2, item 2).
    colunas = [("Nº Licença", 0.11), ("Empresa", 0.20), ("Fase", 0.20), ("Emissão", 0.09),
              ("Validade", 0.10), ("Órgão", 0.10), ("Status", 0.11), ("Vence em", 0.09)]
    linhas_txt = [
        (l["n_licenca"] or "-", l["empresa"] or "", l["fase"] or "", _fmt_data(l["emissao"]),
        _fmt_data(l["validade"]), l["orgao"] or "", l["status"] or "", _fmt_dias(l["vence_em"]))
        for l in linhas_obra
    ]
    img = _desenhar_tabela("Licenças ambientais — %s" % rodovia, colunas,
                           linhas_txt, largura_mm, altura_mm)
    caminho = salvar_png(img, out_dir, "figura6_indices_ambientais.png")
    prov.update({
        "method": ("PIL: tabela nativa das licencas cuja 'Empresa' esta associada a rodovia "
                  "'%s' na aba 'Empresas' (SHADOW_ERROR da v0.1 corrigido: filtro por "
                  "Rodovia, nao por empresa executora do contrato -- ver docstring do "
                  "modulo); NAO reproduz o indice percentual do dashboard Power BI do "
                  "golden (formula nao documentada) -- e' um resumo real e verificavel, "
                  "nao o mesmo dado do golden" % rodovia),
        "road_specific_metrics": {k: v for k, v in metrica_obra.items() if k != "linhas"},
        "corporate_kpis": corporate_kpis,
        "n_licencas": len(linhas_obra), "filtro": "Rodovia == '%s' (via aba Empresas)" % rodovia,
    })
    return FiguraResult(caminho, largura_mm, altura_mm, "OK",
                        "painel de licencas filtrado por rodovia (recorte confirmado na "
                        "Tarefa 1 como o mesmo do relatorio oficial); nao e o indice "
                        "percentual do golden -- ver provenance.method", prov)


def gerar_condicionantes(caminho_licenca_ambiental, rodovia, out_dir) -> FiguraResult:
    largura_mm, altura_mm = 160.0, 63.98  # medida do GOLDEN_VISUAL_SPEC.md Sec.6 (Figura 7)
    prov = {"source_file": caminho_licenca_ambiental, "sheet": "Condicionantes", "rodovia": rodovia}
    if not caminho_licenca_ambiental or not U.exists(caminho_licenca_ambiental):
        return FiguraResult(None, largura_mm, altura_mm, "LIMITATION",
                            "arquivo de licenca ambiental nao encontrado", prov)
    if not rodovia:
        return FiguraResult(None, largura_mm, altura_mm, "LIMITATION",
                            "rodovia nao informada (esperado road.rodovia do report_data)", prov)
    wb = _abrir(caminho_licenca_ambiental)
    por_rodovia_empresas, erro = _ler_empresas_por_rodovia(wb)
    if por_rodovia_empresas is None:
        return FiguraResult(None, largura_mm, altura_mm, "LIMITATION", erro, prov)
    linhas, erro = _ler_condicionantes_todas(wb)
    if linhas is None:
        return FiguraResult(None, largura_mm, altura_mm, "LIMITATION", erro, prov)

    hoje = dt.date.today()
    corporate_kpis, road_specific_metrics = _kpis_condicionantes(linhas, por_rodovia_empresas, hoje)
    metrica_obra = road_specific_metrics.get(_norm_rodovia(rodovia))
    if not metrica_obra or not metrica_obra["linhas"]:
        return FiguraResult(None, largura_mm, altura_mm, "LIMITATION",
                            "nenhuma condicionante resolvida para a rodovia '%s'" % rodovia, prov)
    total = metrica_obra["total_condicionantes"]
    contagem = metrica_obra["contagem_por_status"]

    # Taxonomia com 3+ valores (Tarefa 2, item 3): a ordem e' fixa
    # (_ORDEM_STATUS_CONDICIONANTE), mas NENHUMA categoria e' descartada --
    # o que aparecer nos dados (mesmo fora do vocabulario esperado) entra no
    # final da tabela, nunca e' somado/escondido dentro de outra.
    def _chave_ordem(status):
        try:
            return (0, _ORDEM_STATUS_CONDICIONANTE.index(status))
        except ValueError:
            return (1, status)

    colunas = [("Status", 0.45), ("Quantidade", 0.20), ("% do total (rodovia)", 0.35)]
    linhas_txt = [(status, str(n), "%.1f%%" % (100.0 * n / total) if total else "-")
                 for status, n in sorted(contagem.items(), key=lambda kv: _chave_ordem(kv[0]))]
    linhas_txt.append(("TOTAL — %s" % rodovia, str(total), "100%"))
    img = _desenhar_tabela("Condicionantes ambientais — %s" % rodovia, colunas,
                           linhas_txt, largura_mm, altura_mm)
    caminho = salvar_png(img, out_dir, "figura7_condicionantes.png")
    prov.update({
        "method": ("PIL: contagem por status das condicionantes cuja 'Responsável' esta "
                  "associada a rodovia '%s' na aba 'Empresas' (SHADOW_ERROR da v0.1 "
                  "corrigido: filtro por Rodovia, nao empresa; status recalculado com a "
                  "formula documentada de Condicionantes!G usando a data real de hoje, "
                  "%s -- nao o cache frio da planilha, nao uma data escolhida a dedo, ver "
                  "docstring do modulo) (%d linhas reais); resumido por status (nao item a "
                  "item) porque o volume real nao cabe no painel pequeno do golden "
                  "(160x64mm) -- decisao de design registrada aqui" %
                  (rodovia, hoje.isoformat(), total)),
        "hoje_usado_no_recalculo": hoje.isoformat(),
        "road_specific_metrics": {k: v for k, v in metrica_obra.items() if k != "linhas"},
        "corporate_kpis": corporate_kpis,
        "n_condicionantes": total, "filtro": "Rodovia == '%s' (via aba Empresas)" % rodovia,
    })
    return FiguraResult(caminho, largura_mm, altura_mm, "OK",
                        "resumo por status (%d condicionantes reais, %d categorias -- "
                        "taxonomia completa, nao reduzida a 2); numeros sao o estado ATUAL "
                        "da copia local da planilha (SOURCE_LIMITATION igual ao TODAY() das "
                        "NC, ver docstring do modulo), nao uma reconstrucao do golden -- "
                        "ver provenance.method" % (total, len(contagem)), prov)


def gerar_qrcode(out_dir) -> FiguraResult:
    img, draw = novo_canvas(QR_LADO_MM, QR_LADO_MM)
    w, h = img.size
    desenhar_borda(draw, (0, 0, w - 1, h - 1), mm_to_px(0.6), COR_PRETO)
    margem = mm_to_px(4)
    draw.line([(margem, margem), (w - margem, h - margem)], fill=COR_BORDA, width=mm_to_px(0.4))
    draw.line([(w - margem, margem), (margem, h - margem)], fill=COR_BORDA, width=mm_to_px(0.4))
    fonte = get_font(11, bold=True)
    fonte_nota = get_font(6.5)
    txt = "QR CODE"
    tw = text_width(draw, txt, fonte)
    draw.rectangle([w / 2 - tw / 2 - mm_to_px(2), h / 2 - mm_to_px(6),
                   w / 2 + tw / 2 + mm_to_px(2), h / 2 + mm_to_px(6)], fill=(255, 255, 255))
    draw.text((w / 2 - tw / 2, h / 2 - mm_to_px(3)), txt, font=fonte, fill=COR_PRETO)
    nota = "biblioteca indisponível"
    tw2 = text_width(draw, nota, fonte_nota)
    draw.text((w / 2 - tw2 / 2, h - mm_to_px(7)), nota, font=fonte_nota, fill=COR_BORDA)

    caminho = salvar_png(img, out_dir, "figura8_qrcode_placeholder.png")
    prov = {
        "source_file": None,
        "method": ("placeholder desenhado com PIL; a biblioteca 'qrcode' nao existe neste "
                  "ambiente e nao pode ser instalada (spec Sec.0.1) -- nao se fabrica um "
                  "QR Code falso. Elemento incluido no relatorio, marcado como pendencia "
                  "de apresentacao."),
    }
    return FiguraResult(caminho, QR_LADO_MM, QR_LADO_MM, "PRESENTATION_DEPENDENCY",
                        "QR Code real requer a biblioteca 'qrcode', indisponivel e nao "
                        "instalavel neste ambiente; emitido placeholder marcado", prov)
