# -*- coding: utf-8 -*-
r"""
06_inventariar_ecossistema.py

Inventaria o ECOSSISTEMA DOCUMENTAL que alimenta o Relatorio Semanal de
Acompanhamento de Obras do Programa, nas 4 obras (RD101, RD102A, RD103, RD104).

Escopo varrido
--------------
A) Pastas IRMAS da "5. Relatorio Semanal de Acompanhamento de Obras", dentro de
   "<OBRA>\3_Gestao Tecnica de Obra\4. Relatorio Semanal (de) Obras\":
     1. Ficha(s) de campo
     2. Ficha (de) Acompanhamento Semanal (Obra)
     3. Fotos
     4. Diario(s) de Obra(s)
     (+ "Memorial de campo", quando existir)
B) Demais pastas de "3_Gestao Tecnica de Obra" (Eventograma, Boletins de
   Medicao, Notificacoes, Relatorio Mensal, Ambiental/SST, Documentos
   Recebidos, Atas) -- somente AGREGADO, sem descer arquivo a arquivo.

Politica de volume
------------------
"3. Fotos" tem ~35 mil imagens somadas. Registrar uma linha por foto tornaria o
inventario inutilizavel. Portanto:
  - Cada pasta-dia de "3. Fotos" (e cada subpasta de servico dentro dela) vira
    UMA linha do tipo `foto_lote`, com a contagem de arquivos e os extremos de
    mtime. As fotos individuais NAO viram linha.
  - Todos os demais arquivos (fichas, planilhas, diarios, PDFs) viram uma linha
    cada.

REGRAS DO PROJETO respeitadas aqui:
  - O SharePoint e READ-ONLY: este script so LE de la. Toda escrita vai para
    output/inventarios/ do repositorio, protegida por U.assert_read_only().
  - Todo acesso a arquivo passa por prog_util (prefixo \\?\ para MAX_PATH).

Saidas
------
  output/inventarios/ecossistema.json
  output/inventarios/ecossistema.csv

Uso:  python scripts/06_inventariar_ecossistema.py
"""
import os
import re
import sys
import csv
import json
import datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import prog_util as U  # noqa: E402

# --------------------------------------------------------------------------
# Semanas da amostra (regime vigente). Sabado -> sexta.
# --------------------------------------------------------------------------
SEMANAS_AMOSTRA = [
    ("2026-08-22", "2026-08-28"),
    ("2026-08-29", "2026-09-04"),
]
# Janela alargada: pega o insumo da semana anterior e o arquivo salvo depois.
JANELA_INI = datetime.date(2026, 8, 15)
JANELA_FIM = datetime.date(2026, 9, 16)

EXT_IGNORADAS = {".ini", ".db", ".lnk"}

# --------------------------------------------------------------------------
# Classificacao de funcao provavel (heuristica por pasta + nome + extensao).
# NOME DE ARQUIVO NAO E PROVA DE FUNCAO -- o campo abaixo e uma HIPOTESE de
# trabalho; a verificacao por leitura do conteudo esta em
# docs/mapeamento/06_fontes_dados.md, com nivel de confianca por afirmacao.
# --------------------------------------------------------------------------
FUNCAO_POR_IRMA = {
    "1": "ficha_campo_diaria",
    "2": "ficha_acompanhamento_semanal",
    "3": "acervo_fotografico",
    "4": "diario_de_obra",
}


def _norm(s):
    return U.norm(s)


def classificar(irma_num, rel, nome, ext):
    """Devolve (funcao_provavel, secao_alvo_hipotese)."""
    n = _norm(nome)
    r = _norm(rel)
    if irma_num == "1":
        if "acumulada" in n or "avanco fisico" in n:
            return "ficha_campo_acumulada", "2 Ficha de Acompanhamento / 3 Avanco Global"
        return "ficha_campo_diaria", "5.1 Condicoes climaticas / 5.2 indice de fotos"
    if irma_num == "2":
        return "ficha_acompanhamento_semanal", "2 Ficha / 5.1 Clima / 8 NC / 9 Riscos"
    if irma_num == "3":
        return "acervo_fotografico", "5.2 Relatorio Fotografico por servico"
    if irma_num == "4":
        if "digital" in r:
            return "diario_obra_digital_assinado_semanal", "7 Diario de Obras / 11 Anexo II"
        if "fisico" in r:
            return "diario_obra_fisico_mensal", "7 Diario de Obras (historico)"
        if "obsoleto" in r:
            return "diario_obra_obsoleto", "-"
        return "diario_de_obra", "7 Diario de Obras"
    if irma_num == "M":
        return "memoria_calculo_servicos", "Boletim de Medicao (nao o semanal)"
    return "desconhecido", "-"


# --------------------------------------------------------------------------
# Datas embutidas no nome / caminho
# --------------------------------------------------------------------------
_PAT_YMD = re.compile(r"(20\d\d)[-._](\d{1,2})[-._](\d{1,2})")
_PAT_DMY4 = re.compile(r"(\d{2})[-._](\d{2})[-._](20\d\d)")
_PAT_DMY2 = re.compile(r"(?<!\d)(\d{2})[-._](\d{2})[-._](\d{2})(?!\d)")
_PAT_PASTA_FOTO = re.compile(r"^(\d{2})[-_](\d{2})[-_](\d{2})$")


# As obras comecaram em set/2025. Datas fora desta faixa em nome de arquivo sao
# quase sempre falso positivo do parser (ex.: '24-1.pdf', '25-10-20 a 25-10-27'
# lido como dd-mm-aa). Descartamos em vez de contaminar o inventario.
DATA_MIN = datetime.date(2025, 1, 1)
DATA_MAX = datetime.date(2027, 12, 31)
# Nenhum insumo semanal legitimo cobre mais que ~45 dias; acima disso o
# intervalo e ruido de parsing e o periodo fica em branco.
SPAN_MAX_DIAS = 45


def _d(y, m, d):
    try:
        dt = datetime.date(y, m, d)
    except ValueError:
        return None
    return dt if DATA_MIN <= dt <= DATA_MAX else None


def datas_no_texto(txt):
    """Extrai todas as datas plausiveis de um nome de arquivo/pasta.

    Uma pasta-dia de '3. Fotos' ('26-08-29', '25_10_01') e AA_MM_DD e casaria
    tambem com dd-mm-aa, produzindo uma data espuria em 2029. Por isso esse
    formato e tratado antes, em exclusividade.
    """
    mp = _PAT_PASTA_FOTO.match(txt.strip())
    if mp:
        a, b, c = (int(x) for x in mp.groups())
        d = _d(2000 + a, b, c)
        return [d] if d else []
    out = []
    for m in _PAT_YMD.finditer(txt):
        out.append(_d(int(m.group(1)), int(m.group(2)), int(m.group(3))))
    for m in _PAT_DMY4.finditer(txt):
        out.append(_d(int(m.group(3)), int(m.group(2)), int(m.group(1))))
    for m in _PAT_DMY2.finditer(txt):
        a, b, c = (int(x) for x in m.groups())
        # dd-mm-aa  (ex.: 22-08-26)
        out.append(_d(2000 + c, b, a))
    return [x for x in out if x]


def periodo_do_nome(txt):
    """Devolve (data_ini, data_fim) se o nome codificar um intervalo; senao
    (data, data) se codificar um dia; senao (None, None)."""
    ds = sorted(set(datas_no_texto(txt)))
    if not ds:
        return None, None
    if len(ds) >= 2:
        if (ds[-1] - ds[0]).days > SPAN_MAX_DIAS:
            return None, None
        return ds[0], ds[-1]
    return ds[0], ds[0]


def semanas_cobertas(ini, fim):
    """Quais semanas da amostra o intervalo [ini, fim] toca."""
    if not ini:
        return []
    fim = fim or ini
    out = []
    for a, b in SEMANAS_AMOSTRA:
        sa = datetime.date.fromisoformat(a)
        sb = datetime.date.fromisoformat(b)
        if ini <= sb and fim >= sa:
            out.append(a + ".." + b)
    return out


# --------------------------------------------------------------------------
# Localizacao das pastas irmas
# --------------------------------------------------------------------------
def irmas_da_obra(obra):
    """Devolve lista de (numero_logico, nome_real, caminho) das pastas irmas."""
    raiz = U.pasta_semanal_raiz(obra)
    if not raiz:
        return raiz, []
    out = []
    for nome in U.listdir(raiz):
        p = os.path.join(raiz, nome)
        if not U.isdir(p):
            continue
        if nome.startswith("5. "):
            continue  # e a pasta do proprio relatorio
        m = re.match(r"^(\d)\.", nome)
        if m:
            out.append((m.group(1), nome, p))
        elif "memorial" in _norm(nome):
            out.append(("M", nome, p))
        else:
            out.append(("?", nome, p))
    return raiz, out


# --------------------------------------------------------------------------
# Varredura
# --------------------------------------------------------------------------
def varrer_fotos(obra, irma_nome, base):
    """Uma linha por pasta-dia e por subpasta de servico. Nunca por foto."""
    linhas = []
    for dia in U.listdir(base):
        pdia = os.path.join(base, dia)
        if not U.isdir(pdia):
            continue
        ini, fim = periodo_do_nome(dia)
        sem = semanas_cobertas(ini, fim)
        filhos = U.listdir(pdia)
        servicos = [x for x in filhos if U.isdir(os.path.join(pdia, x))]
        soltos = [x for x in filhos
                  if not U.isdir(os.path.join(pdia, x))
                  and os.path.splitext(x)[1].lower() not in EXT_IGNORADAS]

        def lote(caminho, arquivos, servico):
            if not arquivos:
                return None
            tam = 0
            mts = []
            exts = {}
            for f in arquivos:
                st = U.stat_info(os.path.join(caminho, f))
                tam += st["tamanho_bytes"]
                mts.append(st["modificado"])
                e = os.path.splitext(f)[1].lower()
                exts[e] = exts.get(e, 0) + 1
            return {
                "obra": obra,
                "pasta_irma": irma_nome,
                "tipo": "foto_lote",
                "caminho": caminho,
                "nome": os.path.basename(caminho),
                "ext": "|".join("%s:%d" % kv for kv in sorted(exts.items(), key=lambda x: -x[1])),
                "qtd_arquivos": len(arquivos),
                "tamanho_bytes": tam,
                "modificado_min": min(mts)[:19],
                "modificado_max": max(mts)[:19],
                "dia": dia,
                "servico": servico,
                "periodo_ini": ini.isoformat() if ini else "",
                "periodo_fim": fim.isoformat() if fim else "",
                "semanas_amostra": ";".join(sem),
                "funcao_provavel": "acervo_fotografico",
                "secao_alvo_hipotese": "5.2 Relatorio Fotografico por servico",
            }

        r = lote(pdia, soltos, "")
        if r:
            linhas.append(r)
        for s in servicos:
            ps = os.path.join(pdia, s)
            arqs = [x for x in U.listdir(ps)
                    if not U.isdir(os.path.join(ps, x))
                    and os.path.splitext(x)[1].lower() not in EXT_IGNORADAS]
            r = lote(ps, arqs, s)
            if r:
                linhas.append(r)
    return linhas


def varrer_documental(obra, irma_num, irma_nome, base):
    """Uma linha por arquivo (fichas, planilhas, diarios)."""
    linhas = []
    for dirpath, dirnames, filenames in U.walk(base):
        rel = os.path.relpath(dirpath, base)
        rel = "" if rel == "." else rel
        for f in filenames:
            ext = os.path.splitext(f)[1].lower()
            if ext in EXT_IGNORADAS:
                continue
            p = os.path.join(dirpath, f)
            st = U.stat_info(p)
            ini, fim = periodo_do_nome(f)
            if not ini and rel:
                ini, fim = periodo_do_nome(rel)
            func, secao = classificar(irma_num, rel, f, ext)
            mt = datetime.date.fromisoformat(st["modificado"][:10])
            linhas.append({
                "obra": obra,
                "pasta_irma": irma_nome,
                "tipo": "arquivo",
                "caminho": p,
                "nome": f,
                "subpasta": rel,
                "ext": ext,
                "qtd_arquivos": 1,
                "tamanho_bytes": st["tamanho_bytes"],
                "modificado_min": st["modificado"][:19],
                "modificado_max": st["modificado"][:19],
                "periodo_ini": ini.isoformat() if ini else "",
                "periodo_fim": fim.isoformat() if fim else "",
                "semanas_amostra": ";".join(semanas_cobertas(ini, fim)),
                "na_janela_mtime": JANELA_INI <= mt <= JANELA_FIM,
                "funcao_provavel": func,
                "secao_alvo_hipotese": secao,
            })
    return linhas


def agregar_outras_pastas(obra):
    """Agregado das demais pastas de 3_Gestao Tecnica de Obra (sem descer)."""
    gto = os.path.join(U.BASE_OBRAS, obra, "3_Gestao Tecnica de Obra")
    out = []
    for d in U.listdir(gto):
        p = os.path.join(gto, d)
        if not U.isdir(p) or d.startswith("4. Relat"):
            continue
        n = 0
        tam = 0
        exts = {}
        mts = []
        for dirpath, dirnames, filenames in U.walk(p):
            for f in filenames:
                e = os.path.splitext(f)[1].lower()
                if e in EXT_IGNORADAS:
                    continue
                try:
                    st = U.stat_info(os.path.join(dirpath, f))
                except OSError:
                    continue
                n += 1
                tam += st["tamanho_bytes"]
                mts.append(st["modificado"])
                exts[e] = exts.get(e, 0) + 1
        out.append({
            "obra": obra,
            "pasta": d,
            "caminho": p,
            "qtd_arquivos": n,
            "tamanho_bytes": tam,
            "exts_top": sorted(exts.items(), key=lambda x: -x[1])[:6],
            "subpastas_nivel1": [x for x in U.listdir(p) if U.isdir(os.path.join(p, x))][:12],
            "modificado_min": min(mts)[:19] if mts else "",
            "modificado_max": max(mts)[:19] if mts else "",
        })
    return out


# --------------------------------------------------------------------------
def main():
    destino = os.path.join(U.OUT, "inventarios")
    U.assert_read_only(destino)
    os.makedirs(destino, exist_ok=True)

    inv = {
        "gerado_em": datetime.datetime.now().isoformat(timespec="seconds"),
        "semanas_amostra": ["%s..%s" % s for s in SEMANAS_AMOSTRA],
        "politica_fotos": "pastas de '3. Fotos' agregadas em linhas foto_lote "
                          "(dia e dia/servico); fotos individuais nao viram linha",
        "obras": {},
        "outras_pastas_gto": {},
    }
    todas = []

    for obra in U.OBRAS:
        raiz, irmas = irmas_da_obra(obra)
        print("[%s] raiz4 = %s" % (obra, os.path.basename(raiz or "")))
        bloco = {"raiz_pasta_4": raiz, "pastas_irmas": {}}
        for num, nome, caminho in irmas:
            print("   varrendo %s ..." % nome, flush=True)
            if num == "3":
                linhas = varrer_fotos(obra, nome, caminho)
            else:
                linhas = varrer_documental(obra, num, nome, caminho)
            todas.extend(linhas)
            bloco["pastas_irmas"][nome] = {
                "numero_logico": num,
                "caminho": caminho,
                "n_linhas": len(linhas),
                "n_arquivos": sum(x["qtd_arquivos"] for x in linhas),
                "tamanho_bytes": sum(x["tamanho_bytes"] for x in linhas),
                "linhas_nas_semanas_amostra": [
                    x for x in linhas if x["semanas_amostra"]
                ],
            }
        inv["obras"][obra] = bloco
        print("   agregando demais pastas de 3_Gestao Tecnica de Obra ...", flush=True)
        inv["outras_pastas_gto"][obra] = agregar_outras_pastas(obra)

    inv["total_linhas"] = len(todas)
    inv["linhas"] = todas

    pj = os.path.join(destino, "ecossistema.json")
    U.assert_read_only(pj)
    with open(pj, "w", encoding="utf-8") as fh:
        json.dump(inv, fh, ensure_ascii=False, indent=1, default=str)

    cols = ["obra", "pasta_irma", "tipo", "funcao_provavel", "secao_alvo_hipotese",
            "nome", "subpasta", "dia", "servico", "ext", "qtd_arquivos",
            "tamanho_bytes", "modificado_min", "modificado_max",
            "periodo_ini", "periodo_fim", "semanas_amostra", "caminho"]
    pc = os.path.join(destino, "ecossistema.csv")
    U.assert_read_only(pc)
    with open(pc, "w", encoding="utf-8-sig", newline="") as fh:
        w = csv.DictWriter(fh, fieldnames=cols, extrasaction="ignore", delimiter=";")
        w.writeheader()
        for r in todas:
            w.writerow(r)

    print("\nOK  %d linhas" % len(todas))
    print("    " + pj)
    print("    " + pc)


if __name__ == "__main__":
    main()
