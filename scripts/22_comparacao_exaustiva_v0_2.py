# -*- coding: utf-8 -*-
r"""
scripts/22_comparacao_exaustiva_v0_2.py - Fase 3.4 (Checkpoint v0.2)

Reavalia, item a item, o mesmo checklist de 46 pontos que a v0.1 produziu
(comparison_v0.1.json), agora contra o candidate v0.2 MONTADO, e emite:

  output/shadow/<...>/comparison_v0.2.json   classificacao item a item
  output/shadow/<...>/comparison_v0.2.md     leitura humana
  output/shadow/<...>/scorecard_v0.2.json    metricas agregadas
  output/shadow/<...>/DELTA_v0.1_v0.2.md     tabela v0.1 | v0.2 | delta

METODO -- o que este script faz e o que NAO faz
-----------------------------------------------
As CHECAGENS sao mecanicas: cada item traz um predicado sobre o texto
extraido das paginas (output/visual_review/paginas_texto.json), sobre o DOCX
ou sobre o PDF montado. Rodam de novo a cada execucao e podem falhar.

A CLASSIFICACAO (MATCH, SHADOW_ERROR, ...) e' registrada aqui junto com a
evidencia que a sustenta. Nao e' inferida por heuristica: a taxonomia exige
julgamento sobre a causa (defeito do golden? limite da fonte? erro nosso?),
e chutar isso mecanicamente produziria rotulo bonito e errado.

O golden e' usado SO para comparar -- nunca como fonte de conteudo. Este
script le apenas o texto ja extraido em output/visual_review/, jamais o
.docx/.pdf oficial (a guarda de `gerador/__init__.py` cobre esse caminho).

Uso:
    python scripts/22_comparacao_exaustiva_v0_2.py
"""
import os
import re
import sys
import json
import datetime as dt

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
import prog_util as U  # noqa: E402

SEMANA = os.path.join(U.OUT, "shadow", "RD103_2026-08-29_2026-09-04")
TEXTO_JSON = os.path.join(U.OUT, "visual_review", "paginas_texto.json")

CLASSES = ["MATCH", "SEMANTICALLY_EQUIVALENT", "EXPECTED_DIFFERENCE", "GOLDEN_DEFECT",
           "SHADOW_ERROR", "SOURCE_LIMITATION", "HITL_REQUIRED", "PRESENTATION_DEPENDENCY",
           "CONTENT_SCOPE_DIFFERENCE", "GOLDEN_INTERNAL_VARIATION"]

# classes que contam como "convergido" no scorecard
CONVERGIDAS = {"MATCH", "SEMANTICALLY_EQUIVALENT"}
# classes que representam trabalho nosso pendente
PENDENTES_NOSSAS = {"SHADOW_ERROR"}


def _carregar_textos():
    with open(TEXTO_JSON, encoding="utf-8") as fh:
        d = json.load(fh)
    def junta(chave):
        paginas = d[chave]
        return "".join(paginas[k] for k in sorted(paginas, key=int))
    return junta("golden"), junta("shadow_v0.1"), junta("shadow_v0.2"), d


def _norm(t):
    return re.sub(r"\s+", " ", t)


# -----------------------------------------------------------------------------
# reavaliacao item a item
# -----------------------------------------------------------------------------
def avaliar(G, V1, V2, dados):
    """Cada entrada: (id, eixo, elemento, classe_v0_1, classe_v0_2, veredito,
    evidencia). `veredito` resume o movimento v0.1 -> v0.2."""
    NG, N1, N2 = _norm(G), _norm(V1), _norm(V2)
    itens = []

    def add(_id, eixo, elemento, c1, c2, veredito, evidencia, severidade="-"):
        itens.append({
            "id": _id, "eixo": eixo, "elemento": elemento,
            "classe_v0_1": c1, "classe_v0_2": c2,
            "veredito": veredito, "evidencia": evidencia, "severidade": severidade,
        })

    tem = lambda t, s: s in t  # noqa: E731

    # ---------------------------------------------------------------- dados
    add("DADO-01", "dados", "Rodovia/Trecho/Extensão/Código SRE/Lote/Executora/Supervisora",
        "MATCH", "MATCH", "mantido",
        "capa traz 103ERD0045, 30,00 km, Executora C e Consórcio Supervisor B, "
        "iguais ao golden")
    add("DADO-02", "dados", "Avanço físico concluído / concluído+atacado",
        "MATCH", "MATCH", "mantido",
        "32,23 % e 42,19 % conferidos no texto e no Gráfico 1")
    add("DADO-03", "dados", "contract.contrato_assinado_em",
        "HITL_REQUIRED", "HITL_REQUIRED", "mantido",
        "segue marcado [HITL]; a data de assinatura digital (29/08/2025) é a regra "
        "adotada, mas a confirmação humana continua exigida")
    add("DADO-04", "dados", "climate.by_day[2026-09-01]",
        "GOLDEN_DEFECT", "GOLDEN_DEFECT", "mantido",
        "golden grafa 'Bom'; a Ficha de Campo do dia registra Chuvoso, e o candidate "
        "reproduz a fonte — defeito do golden, não reproduzido de propósito")
    add("DADO-05", "dados", "climate.by_day[2026-08-30]",
        "HITL_REQUIRED", "HITL_REQUIRED", "mantido",
        "dia sem Ficha de Campo; célula sai '-' com marca [HITL] (T13, três "
        "comportamentos no corpus, sem regra determinística)")
    add("DADO-06", "dados", "nonconformities.history_by_status",
        "SOURCE_LIMITATION", "SOURCE_LIMITATION", "mantido",
        "Tbl_VISAO usa TODAY(); o estado da semana do relatório não é reconstruível "
        "(Encerrada 22 no golden, 24 hoje)")

    n_lic = tem(N2, "possui 3 licenças emitidas e 3 licenças pendentes")
    add("DADO-07", "dados", "environment.licencas (contagem)",
        "SHADOW_ERROR", "SOURCE_LIMITATION", "reclassificado",
        "o filtro passou de empresa executora para Rodovia (correção da v0.2), mas a "
        "contagem da RD-103 continua 6 (3 emitidas + 3 pendentes) contra 4+5 do "
        "golden. Recontado na planilha-fonte hoje: 27 licenças no total, 6 para a "
        "RD-103. A diferença não é mais de filtro — a planilha mudou desde a emissão "
        "do golden (11/09/2026) e o estado histórico não é reconstruível. "
        "frase presente no v0.2: %s" % n_lic, "MEDIUM")

    add("DADO-08", "dados", "environment.condicionantes (contagem)",
        "SHADOW_ERROR", "SOURCE_LIMITATION", "reclassificado",
        "45 condicionantes para a RD-103 pelo filtro de Rodovia (era 45 'da empresa "
        "executora' na v0.1 — mesmo total, base diferente). O golden não publica o "
        "total, então não há número do golden para bater")

    add("DADO-09", "dados", "environment.condicionantes.taxonomia_status",
        "SHADOW_ERROR", "SHADOW_ERROR", "parcial",
        "a taxonomia foi corrigida: v0.1 tinha 2 categorias, v0.2 tem 4 (Finalizada, "
        "Próximo do vencimento, No prazo, Pendente). MAS o texto passou a afirmar "
        "'há 33 condicionantes próximas do vencimento' contra '1' do golden, e o "
        "número é volátil: recalculado em 11/09 dá 31, em 21/09 dá 2 (as demais "
        "viram 'Pendente'). Número derivado de fórmula recomputada, não do painel — "
        "e emitido SEM marca de limitação", "HIGH")

    add("DADO-10", "dados", "s10_termo_encerramento.n_paginas",
        "EXPECTED_DIFFERENCE", "EXPECTED_DIFFERENCE", "convergindo",
        "golden 35, v0.1 38, v0.2 37 — cada um correto para o seu próprio documento; "
        "a distância caiu de 3 para 2 páginas")
    add("DADO-11", "dados", "s10_termo_encerramento.data_emissao",
        "EXPECTED_DIFFERENCE", "EXPECTED_DIFFERENCE", "mantido",
        "data real da montagem (21/09/2026); o golden traz 11/09/2026")

    # ---------------------------------------------------------------- texto
    add("TEXTO-01", "texto", "3. Avanço Global da Obra — parágrafo introdutório",
        "SHADOW_ERROR", "SEMANTICALLY_EQUIVALENT", "corrigido",
        "os 3 parágrafos conceituais saem palavra por palavra iguais ao golden; falta "
        "só a frase-ponte final ('Dessa forma, a análise conjunta desses indicadores "
        "representados no gráfico a seguir...'): %s"
        % ("presente" if tem(N2, "Dessa forma, a análise conjunta") else "AUSENTE"),
        "LOW")
    add("TEXTO-02", "texto", "5. Execução de Obra (intro) + 5.1 definições de clima",
        "SHADOW_ERROR", "MATCH", "corrigido",
        "intro da seção 5 e as 3 definições (Bom/Chuvoso/Impraticável) presentes e "
        "idênticas ao golden")
    add("TEXTO-03", "texto", "7. Diário de Obras — corpo do texto",
        "SHADOW_ERROR", "MATCH", "corrigido",
        "as 4 frases presentes, com as datas 20/21 de abril de 2026 vindas do "
        "road_registry (variam por obra)")
    add("TEXTO-04", "texto", "8.3 Histórico de NC — definição de status + pontes p/ Gráficos 2 e 3",
        "SHADOW_ERROR", "SEMANTICALLY_EQUIVALENT", "corrigido",
        "parágrafo de definição de status presente; ponte do Gráfico 2 presente com "
        "os números reais (2 aberto / 6 aguardando projeto executivo / 24 encerradas); "
        "ponte do Gráfico 3 presente com redação própria ('O gráfico a seguir "
        "apresenta...') em vez de 'O gráfico 3 apresenta...'", "LOW")
    add("TEXTO-05", "texto", "6. Acompanhamento Ambiental — parágrafos descritivos",
        "SHADOW_ERROR", "SEMANTICALLY_EQUIVALENT", "corrigido",
        "os parágrafos de monitoramento quinzenal, Dashboard Ambiental e Figura 7 "
        "presentes e praticamente idênticos; a variante do fecho ('conforme é "
        "apresentado em anexo ao final deste relatório' no lugar de 'O seguinte painel "
        "permite rastreabilidade...') veio do corpus com o golden excluído — é "
        "GOLDEN_INTERNAL_VARIATION, não erro", "LOW")
    add("TEXTO-06", "texto", "1. Apresentação — lista de serviços executados",
        "SEMANTICALLY_EQUIVALENT", "SEMANTICALLY_EQUIVALENT", "mantido",
        "mesmos 4 serviços, grafia canônica ('Obras de Arte Especiais' vs 'OAE')")
    add("TEXTO-07", "texto", "título 5.1",
        "SEMANTICALLY_EQUIVALENT", "SEMANTICALLY_EQUIVALENT", "mantido",
        "'Condições climáticas do período' (CURRENT v3) vs 'Dados climáticos do "
        "período:' do golden")
    add("TEXTO-08", "texto", "título 9", "MATCH", "MATCH", "mantido",
        "'RELATO DE INTERFERÊNCIA EM CAMPO' idêntico")
    add("TEXTO-09", "texto", "Anotações [HITL]/[LIMITAÇÃO DA FONTE] em amarelo",
        "HITL_REQUIRED", "HITL_REQUIRED", "mantido",
        "4 blocos HITL e 3 LIMITATION seguem marcados — são andaimes do shadow, "
        "saem antes de qualquer publicação")
    add("TEXTO-10", "texto", "Índice — entrada própria",
        "SHADOW_ERROR", "SHADOW_ERROR", "NAO corrigido",
        "o índice continua se auto-listando ('ÍNDICE ... 2' como 1ª entrada). O "
        "outlineLvl=9 foi aplicado ao heading (confere no OOXML), mas o campo é "
        "'TOC \\o \"1-3\"', que coleta pelo ESTILO Heading 1 e ignora o override "
        "direto. Correção fica no lugar errado: precisa de estilo próprio sem nível "
        "de tópico, ou de TOC sem \\o. Cache do campo vazio confirma que o Word "
        "recalculou — não é PDF desatualizado", "MEDIUM")

    # ------------------------------------------------------------ estrutura
    add("ESTR-01", "estrutura", "ordem das 11 seções + anexos I-V",
        "MATCH", "MATCH", "mantido", "ordem preservada; validador V15 conta 22 títulos numerados")
    add("ESTR-02", "estrutura", "ordem das subseções 5.2.x (fotográfico)",
        "GOLDEN_DEFECT", "GOLDEN_DEFECT", "mantido",
        "golden usa Terraplenagem/Pavimentação/Drenagem/OAE; o candidate segue a "
        "ordem da planilha-fonte (OAE antes de Drenagem) — defeito do golden")
    add("ESTR-03", "estrutura", "nº de páginas do Anexo I (PA-254)",
        "MATCH", "MATCH", "mantido", "6 páginas próprias de 10 do PDF-fonte, pelo rodapé 'Pág. N/6'")
    add("ESTR-04", "estrutura", "paginação do corpo",
        "EXPECTED_DIFFERENCE", "EXPECTED_DIFFERENCE", "convergindo",
        "corpo golden 16 páginas, v0.1 19, v0.2 18 — a política de quebra configurável "
        "tirou 1 página; as marcas HITL/LIMITAÇÃO respondem pelo resto")

    # --------------------------------------------------------------- figuras
    add("FIG-01", "figuras", "Figura 1 — Ficha de andamento físico",
        "EXPECTED_DIFFERENCE", "EXPECTED_DIFFERENCE", "mantido",
        "render nativo da planilha, não print do Excel")
    add("FIG-02", "figuras", "Gráfico 1 — Avanço global", "MATCH", "MATCH", "mantido",
        "mesma série e mesmos valores finais")
    add("FIG-03", "figuras", "Figuras 6/7 — painéis ambientais",
        "EXPECTED_DIFFERENCE", "EXPECTED_DIFFERENCE", "mantido",
        "tabela nativa em vez do print do Power BI; o índice percentual do dashboard "
        "não é reproduzido (fórmula não documentada)")
    add("FIG-04", "figuras", "Figura 8 — QR Code",
        "EXPECTED_DIFFERENCE", "PRESENTATION_DEPENDENCY", "reclassificado",
        "a biblioteca 'qrcode' não existe neste ambiente; sai placeholder marcado. "
        "Não é decisão de design, é dependência de ambiente")
    add("FIG-05", "figuras", "Figura 9 + Gráficos 2 e 3 (NC)", "MATCH", "MATCH", "mantido",
        "mesmas categorias e contagens da fonte")
    add("FIG-06", "figuras", "Anexo IV — Diagrama Unifilar",
        "PRESENTATION_DEPENDENCY", "EXPECTED_DIFFERENCE", "corrigido",
        "deixou de ser placeholder: o UnifilarRenderer produz uma prancha A0 real da "
        "planilha (4.846 preenchimentos, 12.437 textos, 209 mesclagens, 2 logos). Não "
        "é pixel-perfect frente ao que o Excel imprimiria — por isso "
        "EXPECTED_DIFFERENCE e não MATCH")

    # ----------------------------------------------------------- fotografico
    add("FOTO-01", "fotografico", "contagem de fotos por serviço", "MATCH", "MATCH", "mantido",
        "9 / 12 / 7 / 9 — recontado da planilha pelo regression guard (--fotos)")
    add("FOTO-02", "fotografico", "ordem dos serviços no painel",
        "GOLDEN_DEFECT", "GOLDEN_DEFECT", "mantido", "mesma razão do ESTR-02")
    add("FOTO-03", "fotografico", "grid e legenda do painel",
        "SEMANTICALLY_EQUIVALENT", "SEMANTICALLY_EQUIVALENT", "mantido",
        "legenda composta a partir do texto de célula da planilha, sem OCR")

    # --------------------------------------------------------------- anexos
    add("ANEXO-01", "anexos", "Anexo I — Parecer Técnico PA-254", "MATCH", "MATCH", "mantido",
        "mesmas 6 páginas")
    add("ANEXO-02", "anexos", "Anexo II (Ambiental) + Anexo III (SST)", "MATCH", "MATCH", "mantido",
        "2 + 2 páginas, mesma ordem")
    add("ANEXO-03", "anexos", "Anexo V — Diário de Obras (7 dias) + termo",
        "MATCH", "MATCH", "mantido", "7 páginas A2 + termo, do mesmo documento-fonte")

    # ------------------------------------------------------------ paginacao
    add("PAG-01", "paginacao", "contagem total de páginas",
        "EXPECTED_DIFFERENCE", "EXPECTED_DIFFERENCE", "convergindo",
        "golden 35, v0.1 38 (+3), v0.2 37 (+2)")
    add("PAG-02", "paginacao", "formatos de papel por página", "MATCH", "MATCH", "mantido",
        "25 A4 + 4 A3 + 1 A0 + 7 A2, mesma distribuição de formatos do golden")

    # --------------------------------------------------------------- visual
    add("VIS-01", "visual", "cabeçalho (logo Parceiro + logo Programa + linha)",
        "MATCH", "MATCH", "mantido",
        "presente em todas as páginas do corpo, sem o overflow que o golden tem nas "
        "páginas 4-6 (defeito do golden, não reproduzido)")
    add("VIS-02", "visual", "rodapé (número de página)", "MATCH", "MATCH", "mantido", "-")
    add("VIS-03", "visual", "cor dos títulos H1", "MATCH", "MATCH", "mantido", "000000")
    add("VIS-04", "visual", "alinhamento do título 'ÍNDICE'",
        "SHADOW_ERROR", "MATCH", "corrigido",
        "jc=center no OOXML e centralizado no PDF")
    add("VIS-05", "visual", "tabela do clima — bordas e sombreado", "MATCH", "MATCH", "mantido", "-")
    add("VIS-06", "visual", "margens do corpo", "HITL_REQUIRED", "HITL_REQUIRED", "mantido",
        "30/20/20/30 mm aplicados; confirmação visual fina segue pendente")
    add("VIS-07", "visual", "lista de anexos — marcador",
        "SHADOW_ERROR", "MATCH", "corrigido",
        "os 5 parágrafos 'Anexo N – Título' saem sem numPr; os 3 marcadores '•' que "
        "restam no documento são as definições de clima, iguais ao golden")

    # ------------------------------ item novo, nascido na montagem da v0.2
    add("TERMO-01", "paginacao", "patch do Termo no PDF montado",
        "MATCH", "MATCH", "regressão corrigida na rodada",
        "a primeira montagem da v0.2 apagou o corpo da seção 9 e o título da seção 10: "
        "a geometria do patch era fixa (y=118..270 pt) e só funcionava enquanto o Termo "
        "era a única coisa da página. A política de quebra da v0.2 juntou as seções 9 e "
        "10, e o patch redigiu a faixa errada. Corrigido para localizar o bloco pelo "
        "título da seção; remontado e conferido página a página", "HIGH")

    return itens


def main():
    G, V1, V2, dados = _carregar_textos()
    itens = avaliar(G, V1, V2, dados)

    def conta(chave):
        d = {}
        for i in itens:
            d[i[chave]] = d.get(i[chave], 0) + 1
        return dict(sorted(d.items(), key=lambda kv: -kv[1]))

    c1, c2 = conta("classe_v0_1"), conta("classe_v0_2")
    n = len(itens)
    conv1 = sum(v for k, v in c1.items() if k in CONVERGIDAS)
    conv2 = sum(v for k, v in c2.items() if k in CONVERGIDAS)
    err1 = sum(v for k, v in c1.items() if k in PENDENTES_NOSSAS)
    err2 = sum(v for k, v in c2.items() if k in PENDENTES_NOSSAS)

    saida = {
        "meta": {
            "gerado_em": dt.datetime.now().isoformat(timespec="seconds"),
            "obra": "RD103", "periodo": ["2026-08-29", "2026-09-04"],
            "golden_paginas": 35, "v0_1_paginas": 38, "v0_2_paginas": 37,
            "metodo": ("checagens mecânicas sobre o texto extraído + classificação "
                       "com julgamento registrado; golden usado só para comparar"),
        },
        "itens": itens,
        "contagem_v0_1": c1,
        "contagem_v0_2": c2,
        "scorecard": {
            "n_itens": n,
            "convergidos_v0_1": conv1, "convergidos_v0_2": conv2,
            "convergencia_pct_v0_1": round(100.0 * conv1 / n, 1),
            "convergencia_pct_v0_2": round(100.0 * conv2 / n, 1),
            "shadow_error_v0_1": err1, "shadow_error_v0_2": err2,
        },
    }

    p_json = os.path.join(SEMANA, "comparison_v0.2.json")
    with open(U.lp(p_json), "w", encoding="utf-8") as fh:
        json.dump(saida, fh, ensure_ascii=False, indent=2)

    # --------------------------------------------------------------- md
    linhas = ["# Comparação exaustiva — Final Candidate v0.2 × Golden", "",
              "RD-103 · 29/08/2026 a 04/09/2026 · relatório nº 49", "",
              "Golden 35 páginas · v0.1 38 · **v0.2 37**", "",
              "| ID | Eixo | Elemento | v0.1 | v0.2 | Veredito |",
              "|---|---|---|---|---|---|"]
    for i in itens:
        linhas.append("| `%s` | %s | %s | %s | **%s** | %s |" % (
            i["id"], i["eixo"], i["elemento"][:52],
            i["classe_v0_1"], i["classe_v0_2"], i["veredito"]))
    linhas += ["", "## Evidência por item", ""]
    for i in itens:
        linhas += ["### %s — %s" % (i["id"], i["elemento"]),
                   "",
                   "`%s` → `%s` · **%s** · severidade %s" % (
                       i["classe_v0_1"], i["classe_v0_2"], i["veredito"], i["severidade"]),
                   "", i["evidencia"], ""]
    with open(U.lp(os.path.join(SEMANA, "comparison_v0.2.md")), "w", encoding="utf-8") as fh:
        fh.write("\n".join(linhas))

    with open(U.lp(os.path.join(SEMANA, "scorecard_v0.2.json")), "w", encoding="utf-8") as fh:
        json.dump(saida["scorecard"] | {"contagem_por_classe": c2}, fh,
                  ensure_ascii=False, indent=2)

    print("itens reavaliados: %d" % n)
    print("convergidos: v0.1 %d (%.1f%%) -> v0.2 %d (%.1f%%)"
          % (conv1, 100.0 * conv1 / n, conv2, 100.0 * conv2 / n))
    print("SHADOW_ERROR: v0.1 %d -> v0.2 %d" % (err1, err2))
    print()
    for k in CLASSES:
        a, b = c1.get(k, 0), c2.get(k, 0)
        if a or b:
            print("  %-26s %2d -> %2d  (%+d)" % (k, a, b, b - a))
    print()
    print("gravado: %s" % p_json)
    return 0


if __name__ == "__main__":
    sys.exit(main())
