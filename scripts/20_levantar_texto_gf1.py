# -*- coding: utf-8 -*-
"""
20_levantar_texto_gf1.py - Fase 3.4 (v0.2), GF-1: levanta do corpus ja extraido
o texto-padrao que falta em 5 secoes (TEXTO-01..05 de
output/shadow/RD103_2026-08-29_2026-09-04/comparison_v0.1.json):

    3   Avanco Global (3 paragrafos-conceito)
    5/5.1  Execucao de Obra / Clima (intro + 3 definicoes climaticas)
    6   Ambiental (4 paragrafos descritivos)
    7   Diario de Obras (historico + formalizacao + "encontra-se anexo")
    8.3 Historico de NC (definicao de status + pontes dos Graficos 2/3)

MESMO metodo de scripts/17_levantar_texto_s1.py: le SOMENTE
output/extracao/texto/ (189 relatorios), normaliza quebra de linha do PDF,
agrupa por esqueleto e conta variantes -- o golden desta comparacao
(RD103_RO049.1) fica SEMPRE excluido do levantamento (nunca e fonte de texto,
so de comparacao visual). Onde a redacao varia por OBRA (nao por semana), o
levantamento e' feito por obra tambem -- essa variacao vai para
config/road_registry.yaml, nunca para `if obra ==` em codigo.

Grava output/analises/texto_gf1.json com a contagem bruta de cada variante,
para auditoria -- os textos escolhidos (predominantes) foram transcritos
manualmente para config/textos_v3.yaml com a MESMA contagem aqui reproduzida.

Execute:  python scripts/20_levantar_texto_gf1.py
"""
import os
import re
import sys
import json
import collections

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import prog_util as U

TEXTO_DIR = os.path.join(U.OUT, "extracao", "texto")
GOLDEN = "RD103_RO049.1"


def _arquivos():
    return sorted(f for f in os.listdir(TEXTO_DIR)
                  if f.endswith(".txt") and GOLDEN not in f and "OBSOLETO" not in f.upper())


def _obra_de(nome_arquivo):
    return nome_arquivo.split("_")[0]


def _norm(s):
    return re.sub(r"\s+", " ", s).strip()


def _contar(pattern, corte=None, por_obra=False):
    """Acha `pattern` em cada relatorio (golden excluido), agrupa por
    esqueleto EXATO (apos normalizar espaco em branco e cortar em `corte`, se
    dado). Devolve (total_com_secao, Counter[esqueleto] ou dict[obra]->Counter,
    exemplos {esqueleto: nome_arquivo})."""
    total = 0
    exemplos = {}
    geral = collections.Counter()
    por_obra_cnt = collections.defaultdict(collections.Counter)
    for f in _arquivos():
        caminho = os.path.join(TEXTO_DIR, f)
        t = open(caminho, encoding="utf-8", errors="replace").read()
        m = pattern.search(t)
        if not m:
            continue
        total += 1
        frag = _norm(m.group(0))
        if corte:
            frag = frag.split(corte)[0]
        frag = frag[:900]
        geral[frag] += 1
        por_obra_cnt[_obra_de(f)][frag] += 1
        exemplos.setdefault(frag, f)
    return total, geral, por_obra_cnt, exemplos


def _resumo(nome, pattern, corte=None, top=4):
    total, geral, por_obra_cnt, exemplos = _contar(pattern, corte)
    mais_comuns = geral.most_common(top)
    print("=== %s -- relatorios com a secao: %d, variantes: %d" % (nome, total, len(geral)))
    for frag, n in mais_comuns:
        pct = 100.0 * n / total if total else 0.0
        print("  %3d (%5.1f%%) [%s] %s" % (n, pct, exemplos[frag], frag[:160]))
    print()
    return {
        "relatorios_com_a_secao": total,
        "variantes": len(geral),
        "top": [{"n": n, "pct": round(100.0 * n / total, 1) if total else 0.0,
                "exemplo_arquivo": exemplos[frag], "texto": frag}
               for frag, n in mais_comuns],
    }


def _resumo_por_obra(nome, pattern, classificador, top=4):
    """Como `_resumo`, mas quebrado por obra -- usado onde a redacao varia
    por obra (licencas, condicionantes, definicao de status, Grafico 3)."""
    total, geral, por_obra_cnt, exemplos = _contar(pattern)
    print("=== %s (por obra)" % nome)
    saida = {}
    for obra in sorted(por_obra_cnt):
        cnt = por_obra_cnt[obra]
        n_obra = sum(cnt.values())
        classes = collections.Counter()
        for frag, n in cnt.items():
            classes[classificador(frag)] += n
        print("  %s: %d ocorrencias -- %s" % (obra, n_obra, dict(classes)))
        saida[obra] = {"total": n_obra, "classes": dict(classes)}
    print()
    return saida


def main():
    out = {"golden_excluido": GOLDEN, "metodo": "scripts/17_levantar_texto_s1.py, generalizado"}

    # ---------------------------------------------------------------- Secao 3
    out["s3_avanco"] = _resumo(
        "Secao 3 -- paragrafo 1 (conceito de avanco global)",
        re.compile(r"O avanço global da obra é.{0,1500}", re.S),
        corte="4. UNIFILAR")

    # ---------------------------------------------------------------- Secao 5/5.1
    out["s5_1_intro"] = _resumo(
        "Secao 5.1 -- intro + definicoes climaticas",
        re.compile(r"Este item visa apresentar.{0,700}", re.S),
        corte="Segue abaixo")

    # ---------------------------------------------------------------- Secao 6
    out["s6_para_a_monitoramento"] = _resumo(
        "Secao 6 -- paragrafo A (monitoramento quinzenal)",
        re.compile(r"O monitoramento do acompanhamento ambiental.{0,500}", re.S),
        corte="O Dashboard")
    out["s6_para_b_dashboard"] = _resumo(
        "Secao 6 -- paragrafo B (Dashboard Ambiental)",
        re.compile(r"O Dashboard Ambiental apresenta.{0,600}", re.S),
        corte="No período de refer")

    def _classificar_licencas(frag):
        if "atos autorizativos ambientais válidos" in frag:
            return "modelo_A"
        if "sendo" in frag and "licenças emitidas" in frag and "solicitadas" in frag:
            return "modelo_B"
        if "solicitadas" in frag and "status de atendimento" in frag:
            return "modelo_D"
        if "licenças emitidas e" in frag:
            return "modelo_C"
        return "outro"

    out["s6_para_c_licencas_por_obra"] = _resumo_por_obra(
        "Secao 6 -- paragrafo C (contagem de licencas), por obra",
        re.compile(r"No período de referência, o empreendimento possui.{0,300}", re.S),
        _classificar_licencas)

    def _classificar_condicionante(frag):
        return "negativa" if "não há condicionantes" in frag else "positiva"

    out["s6_para_d_condicionantes_por_obra"] = _resumo_por_obra(
        "Secao 6 -- paragrafo D (condicionantes / Figura 7), por obra",
        re.compile(r"O painel da figura.{0,400}", re.S),
        _classificar_condicionante)

    # ---------------------------------------------------------------- Secao 7
    out["s7_historico"] = _resumo(
        "Secao 7 -- historico do formato de registro",
        re.compile(r"O Diário de Obras passou por uma atualização.{0,420}", re.S))
    out["s7_formalizacao"] = _resumo(
        "Secao 7 -- formalizacao/assinatura (2a frase do mesmo paragrafo)",
        re.compile(r"Esse documento é submetido.{0,320}", re.S))

    def _classificar_anexo(frag):
        if "e pelo relat" in frag or "semana passada" in frag or "semana anterior" in frag:
            return "multi_periodo (fato, nao forma)"
        if ", encontra-se anexo" in frag or ", encontra-se anexos" in frag:
            return "com_virgula"
        return "sem_virgula"

    out["s7_diario_anexo_por_forma"] = _resumo_por_obra(
        "Secao 7 -- 'encontra-se anexo' (forma com/sem virgula)",
        re.compile(r"O Diário de Obras correspondente.{0,900}", re.S),
        _classificar_anexo)

    # ---------------------------------------------------------------- Secao 8.3
    def _classificar_status_def(frag):
        return "generico" if "a notificação encaminhada ao Parceiro foi devidamente" in frag else "com_numero_embutido"

    out["s8_3_status_definicao_por_obra"] = _resumo_por_obra(
        "Secao 8.3 -- definicao de status, por obra (generico x com numero embutido)",
        re.compile(r"O status e a classificação do histórico.{0,700}", re.S),
        _classificar_status_def)

    out["s8_3_grafico2_ponte"] = _resumo(
        "Secao 8.3 -- ponte do Grafico 2 (status)",
        re.compile(r"O Gráfico \d+ apresenta a distribuição de não conformidades por status.{0,300}", re.S))

    def _classificar_grafico3(frag):
        return "a_seguir" if frag.lower().startswith("o gráfico a seguir") else "numero"

    out["s8_3_grafico3_ponte_por_obra"] = _resumo_por_obra(
        "Secao 8.3 -- ponte do Grafico 3 (natureza), por obra ('O Grafico N' x 'a seguir')",
        re.compile(r"(O Gráfico \d+|O gráfico a seguir) apresenta a distribuição das não conformidades classificadas.{0,220}", re.S),
        _classificar_grafico3)

    dest = os.path.join(U.OUT, "analises", "texto_gf1.json")
    U.assert_read_only(dest)
    os.makedirs(os.path.dirname(dest), exist_ok=True)
    with open(dest, "w", encoding="utf-8") as fh:
        json.dump(out, fh, ensure_ascii=False, indent=1)
    print("json:", dest)


if __name__ == "__main__":
    main()
