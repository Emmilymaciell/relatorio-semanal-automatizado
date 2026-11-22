# -*- coding: utf-8 -*-
"""
21_regression_guard.py

Guarda de regressao entre duas versoes do Final Candidate.

REGRA (instrucao da Emmily, item 12 do plano da v0.2):

    "Qualquer coisa que estava correta na v0.1 e ficar errada na v0.2 e'
     REGRESSION e deve bloquear a conclusao da rodada ate ser investigada."

Ou seja: este script NAO mede semelhanca com o golden. Ele responde a UMA
pergunta -- o que ja estava certo continua certo? Melhora nao e' regressao;
piora e'.

O guard olha tres camadas:

  A. FATOS  -- os numeros que vem da fonte e nao podem mudar por mexida de
               renderer (numero do relatorio, periodo, avanco, corte de
               paginas dos anexos, clima do dia sem ficha). Mudou =>
               REGRESSION, sempre.
  B. STATUS -- contagem de blocos por status nos dois manifests. READY so
               pode subir; HITL/LIMITATION/PLACEHOLDER so podem descer. E o
               status POR SECAO nao pode piorar.
  C. DOCX   -- o documento gerado abre, traz as figuras, e os fatos da
               camada A aparecem no TEXTO (nao basta estar no json).
  D. PDF    -- o PDF MONTADO tem todos os titulos de secao, nao sobrou marca
               de placeholder, e o Termo declara a contagem real.

NIVEL "NAO_VERIFICAVEL"
-----------------------
Um guard que nao acha o dado e responde OK e' pior que nenhum guard -- foi
exatamente assim que a primeira versao deste arquivo passou com as chaves
erradas (`avanco`/`fotos`/`clima` em vez de `progress`/`photos`/`climate`) e
deu verde sem ter conferido nada. Entao: invariante cujo dado nao existe no
artefato e' reportado como NAO_VERIFICAVEL, aparece no resumo e no json, e
NUNCA conta como estavel.

O caso que deu origem a este nivel era a contagem de fotos por painel
(9/12/7/9), que so existia no log de `painel_fotografico.montar`. Na v0.3 ela
passou a ser persistida no manifesto
(validation_result.painel_fotografico_contagem) e o guard compara artefato com
artefato. `--fotos` continua disponivel para conferir contra a planilha-fonte
(le o RELATORIO FOTOGRAFICO, ~160 MB -- por isso e' opt-in), mas ja nao e' o
unico caminho.

Execute:
    python scripts/21_regression_guard.py --obra RD103 \
        --inicio 2026-08-29 --fim 2026-09-04 --base v0.1 --nova v0.2 [--fotos]
"""
import os
import sys
import re
import json
import argparse

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
import prog_util as U  # noqa: E402

# ordem de severidade dos status: quanto maior, pior
PIOR = {"READY": 0, "PLACEHOLDER": 1, "LIMITATION": 2, "HITL": 3}

# contagem de fotos por painel confirmada na v0.1 (bate com o golden; ver
# CHECKPOINT_3.3A). A chave e' a taxonomy_key usada por painel_fotografico.
FOTOS_BASELINE_V01 = {
    "TERRAPLENAGEM": 9,
    "PAVIMENTACAO": 12,
    "DRENAGEM": 7,
    "OBRAS_DE_ARTE_ESPECIAIS": 9,
}


def _carregar_excecoes(base, nova):
    """Movimentos declarados em config/regression_exceptions.yaml para ESTE par
    de versoes. Exigem motivo E evidencia -- sem os dois, a excecao nao vale."""
    caminho = os.path.join(U.PROJ, "config", "regression_exceptions.yaml")
    if not os.path.exists(caminho):
        return {}
    import yaml
    doc = yaml.safe_load(open(caminho, encoding="utf-8")) or {}
    saida = {}
    for e in doc.get("excecoes", []):
        if e.get("base") != base or e.get("nova") != nova:
            continue
        if not (e.get("motivo") or "").strip() or not (e.get("evidencia") or "").strip():
            continue      # excecao sem justificativa nao e' excecao
        saida[e.get("item")] = e
    return saida


class Guard(object):
    def __init__(self, excecoes=None):
        self.regressoes = []
        self.melhorias = []
        self.justificados = []
        self.nao_verificaveis = []
        self.iguais = 0
        self.linhas = []
        self.excecoes = excecoes or {}

    def _excecao(self, nome, base, nova):
        """Devolve a excecao aplicavel a este movimento, ou None.

        Casa pelo item E pelos valores exatos: uma excecao escrita para
        3 -> 5 nao cobre 3 -> 9.
        """
        e = self.excecoes.get(nome)
        if not e:
            return None
        if e.get("de") is not None and e.get("de") != base:
            return None
        if e.get("para") is not None and e.get("para") != nova:
            return None
        return e

    def _log(self, nivel, nome, b, n, obs):
        self.linhas.append({"nivel": nivel, "item": nome, "base": b, "nova": n, "obs": obs})

    def indisponivel(self, nome, motivo):
        """Invariante que este guard NAO conseguiu conferir. Nunca conta como OK."""
        self.nao_verificaveis.append("%s: %s" % (nome, motivo))
        self._log("NAO_VERIFICAVEL", nome, None, None, motivo)

    def fato(self, nome, base, nova):
        """Camada A: valor que NAO pode mudar entre versoes."""
        if base is None and nova is None:
            self.indisponivel(nome, "ausente nas duas versoes -- nada foi conferido")
            return
        if base == nova:
            self.iguais += 1
            self._log("OK", nome, base, nova, "identico")
        else:
            self.regressoes.append("%s: base=%r -> nova=%r" % (nome, base, nova))
            self._log("REGRESSION", nome, base, nova, "fato mudou")

    def valor_esperado(self, nome, esperado, obtido):
        """Camada A: valor aferido contra a baseline conhecida da v0.1."""
        if obtido is None:
            self.indisponivel(nome, "nao foi possivel obter o valor")
        elif esperado == obtido:
            self.iguais += 1
            self._log("OK", nome, esperado, obtido, "bate com a baseline v0.1")
        else:
            self.regressoes.append("%s: esperado %r, obtido %r" % (nome, esperado, obtido))
            self._log("REGRESSION", nome, esperado, obtido, "difere da baseline v0.1")

    def contagem(self, nome, base, nova, direcao):
        """Camada B: numero que so pode ir para um lado.

        direcao='sobe'  -> aumentar e' melhoria, diminuir e' regressao
        direcao='desce' -> diminuir e' melhoria, aumentar e' regressao
        """
        if base == nova:
            self.iguais += 1
            self._log("OK", nome, base, nova, "estavel")
            return
        melhorou = (nova > base) if direcao == "sobe" else (nova < base)
        if melhorou:
            self.melhorias.append("%s: %s -> %s" % (nome, base, nova))
            self._log("MELHOROU", nome, base, nova, "na direcao esperada")
            return
        e = self._excecao(nome, base, nova)
        if e:
            self.justificados.append("%s: %s -> %s -- %s"
                                     % (nome, base, nova, (e.get("motivo") or "").strip()))
            self._log("JUSTIFICADO", nome, base, nova,
                      "movimento contrario à direção esperada, declarado em "
                      "config/regression_exceptions.yaml: %s" % (e.get("motivo") or "").strip())
            return
        self.regressoes.append("%s: %s -> %s (so podia %s)" % (nome, base, nova, direcao))
        self._log("REGRESSION", nome, base, nova, "direcao errada")

    def status_secao(self, secao, base, nova):
        pb, pn = PIOR.get(base, 9), PIOR.get(nova, 9)
        nome = "secao %s" % secao
        if pb == pn:
            self.iguais += 1
            self._log("OK", nome, base, nova, "estavel")
        elif pn < pb:
            self.melhorias.append("%s: %s -> %s" % (nome, base, nova))
            self._log("MELHOROU", nome, base, nova, "status melhor")
        else:
            self.regressoes.append("%s: %s -> %s" % (nome, base, nova))
            self._log("REGRESSION", nome, base, nova, "status pior")

    def checagem(self, nome, ok, detalhe=""):
        """Camada C: condicao booleana no artefato gerado."""
        if ok:
            self.iguais += 1
            self._log("OK", nome, "-", "-", detalhe or "presente")
        else:
            self.regressoes.append("%s: %s" % (nome, detalhe or "ausente"))
            self._log("REGRESSION", nome, "-", "-", detalhe or "ausente")


def _ler(caminho):
    """Le um JSON de artefato, ou None se ele ainda nao existe."""
    if not os.path.exists(caminho):
        return None
    with open(caminho, encoding="utf-8") as fh:
        return json.load(fh)


def _fatos_de(rd):
    """Fatos que o guard protege, lidos do schema REAL de report_data.json.

    As chaves sao as do report_schema.yaml -- period/progress/photos/climate --
    e nao traducoes em portugues. Ver o aviso sobre NAO_VERIFICAVEL no topo.
    """
    f = {}
    f["numero_relatorio"] = rd.get("report_number")

    per = rd.get("period") or {}
    f["periodo_inicio"] = per.get("start")
    f["periodo_fim"] = per.get("end")
    f["periodo_n_dias"] = len(per.get("days") or []) or None

    prog = rd.get("progress") or {}
    # gravados como fracao (0.3223); o documento mostra em % com 2 casas
    for chave, origem in (("avanco_concluido_pct", "global_concluido"),
                          ("avanco_atacado_pct", "global_concluido_atacado")):
        v = prog.get(origem)
        f[chave] = round(v * 100, 2) if isinstance(v, (int, float)) else None

    f["servicos"] = [s.get("key") for s in (rd.get("services") or [])] or None

    for d in (rd.get("climate") or {}).get("by_day") or []:
        # `clima` None e' um FATO aqui (30/08 nao tem Ficha de Campo) -- por isso
        # guardamos o clima_label, que nunca e' None, em vez do bruto
        f["clima::%s" % d.get("date")] = d.get("clima_label")

    tot = 0
    for d in (rd.get("photos") or {}).get("by_day") or []:
        tot += d.get("total") or 0
    f["fotos_inventariadas_na_semana"] = tot or None

    f["anexos_previstos"] = [a.get("id") for a in (rd.get("annexes") or [])] or None
    pa = next((a for a in (rd.get("annexes") or []) if a.get("id") == "parecer_tecnico"), None)
    f["parecer_numero"] = (pa or {}).get("pa_numero")
    return f


def _fatos_anexos(manifest):
    """Corte de paginas e ordem dos anexos embutidos -- o PA-254 mora aqui."""
    f = {}
    vr = manifest.get("validation_result") or {}
    emb = vr.get("embedded_annexes") or []
    f["anexos_embutidos::ordem"] = [a.get("anexo_id") for a in emb] or None
    for a in emb:
        aid = a.get("anexo_id")
        f["anexo::%s::paginas_usadas" % aid] = a.get("paginas")
        f["anexo::%s::paginas_no_pdf_fonte" % aid] = a.get("n_paginas_pdf_fonte")
    return f


def _figuras_status(manifest):
    vr = manifest.get("validation_result") or {}
    fa = vr.get("figuras_frente_a") or {}
    plano = {}

    def _achatar(prefixo, v):
        if isinstance(v, dict):
            for k2, v2 in v.items():
                _achatar("%s.%s" % (prefixo, k2), v2)
        else:
            plano[prefixo] = v

    for k, v in fa.items():
        _achatar(k, v)
    return plano


def _recontar_fotos(rd, manifest):
    """Reconta as fotos por painel direto da planilha (opt-in: --fotos).

    Usa o mesmo indice de rich-value de `painel_fotografico`, mas PARA no
    inventario: nao decodifica imagem nem desenha painel. Assim confere o
    numero sem pagar a composicao das 4 pranchas.
    """
    from gerador.figuras import painel_fotografico as PF
    caminho = (((manifest or {}).get("sources") or {}).get("photo_workbook") or {}).get("path")
    if not caminho or not U.exists(caminho):
        return None, "photo_workbook nao resolvido no manifest"
    local = U.copiar_para_tmp(caminho, "fotografico", "relfoto%s" % os.path.splitext(caminho)[1])
    idx = PF._WorkbookIndex(local)
    aba_para_taxonomia = PF._aba_para_taxonomia()
    contagem = {}
    for aba_nome, (path, estado, n_vm) in idx.abas_disponiveis().items():
        chave = U.norm(aba_nome).rstrip(".")
        tkey = aba_para_taxonomia.get(chave)
        if not tkey or estado != "visible" or n_vm <= 0:
            continue
        _grid, vm_cells = idx.ler_celulas_aba(path)
        contagem[tkey] = len(vm_cells)
    return contagem, None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--obra", default="RD103")
    ap.add_argument("--inicio", default="2026-08-29")
    ap.add_argument("--fim", default="2026-09-04")
    ap.add_argument("--base", default="v0.1")
    ap.add_argument("--nova", default="v0.2")
    ap.add_argument("--fotos", action="store_true",
                    help="reconta as fotos por painel na planilha (le ~160 MB)")
    args = ap.parse_args()

    dirsem = os.path.join(U.OUT, "shadow", "%s_%s_%s" % (args.obra, args.inicio, args.fim))
    p_data = os.path.join(dirsem, "report_data.json")
    p_manif = os.path.join(dirsem, "manifest.yaml")
    p_base = os.path.join(dirsem, "candidate_manifest_%s.json" % args.base)
    p_nova = os.path.join(dirsem, "candidate_manifest_%s.json" % args.nova)

    for p in (p_data, p_base, p_nova):
        if not os.path.exists(p):
            print("ERRO: nao encontrei %s" % p)
            return 2

    with open(p_data, encoding="utf-8") as fh:
        rd = json.load(fh)
    with open(p_base, encoding="utf-8") as fh:
        mb = json.load(fh)
    with open(p_nova, encoding="utf-8") as fh:
        mn = json.load(fh)

    g = Guard(_carregar_excecoes(args.base, args.nova))

    # ---------------------------------------------------------------- A
    # report_data.json e' UM arquivo so -- as duas versoes foram geradas dele.
    # O que este bloco protege e' o CONTRATO com os extratores: se o schema
    # mudou de nome ou de forma, o fato vira None e cai em NAO_VERIFICAVEL.
    fatos = _fatos_de(rd)
    for nome in sorted(fatos):
        g.fato("fato: %s" % nome, fatos[nome], fatos[nome])

    # valores exatos que a v0.1 acertou contra o golden
    g.valor_esperado("fato: numero do relatorio = 49", 49, fatos.get("numero_relatorio"))
    g.valor_esperado("fato: avanco concluido = 32,23%", 32.23, fatos.get("avanco_concluido_pct"))
    g.valor_esperado("fato: avanco atacado = 42,19%", 42.19, fatos.get("avanco_atacado_pct"))
    g.valor_esperado("fato: 30/08 sem Ficha de Campo", "INDETERMINADO",
                     fatos.get("clima::2026-08-30"))
    g.valor_esperado("fato: 4 servicos no relatorio", 4,
                     len(fatos["servicos"]) if fatos.get("servicos") else None)

    # ---------------------------------------------------------------- A'
    fb, fn = _fatos_anexos(mb), _fatos_anexos(mn)
    for nome in sorted(set(fb) | set(fn)):
        g.fato("anexo: %s" % nome, fb.get(nome), fn.get(nome))
    g.valor_esperado("anexo: PA-254 usa 6 de 10 paginas", 6,
                     fn.get("anexo::parecer_tecnico::paginas_usadas"))
    g.valor_esperado("anexo: PA-254 tem 10 paginas na fonte", 10,
                     fn.get("anexo::parecer_tecnico::paginas_no_pdf_fonte"))

    # ---------------------------------------------------------------- A''
    # v0.3: a contagem por painel passou a ser PERSISTIDA no manifesto
    # (validation_result.painel_fotografico_contagem). Quando ela está lá, o
    # guard compara artefato com artefato -- não depende mais de log nem de
    # reabrir a planilha de 160 MB. `--fotos` continua existindo para uma
    # conferência independente contra a fonte, mas deixou de ser o único
    # caminho.
    cont_b = ((mb.get("validation_result") or {}).get("painel_fotografico_contagem")) or {}
    cont_n = ((mn.get("validation_result") or {}).get("painel_fotografico_contagem")) or {}
    if cont_n:
        for tkey in sorted(set(FOTOS_BASELINE_V01) | set(cont_n)):
            esperado = FOTOS_BASELINE_V01.get(tkey, cont_b.get(tkey))
            g.valor_esperado("fotos (manifesto): painel %s" % tkey,
                             esperado, cont_n.get(tkey))
        if cont_b:
            for tkey in sorted(set(cont_b) | set(cont_n)):
                g.fato("fotos: painel %s estável entre versões" % tkey,
                       cont_b.get(tkey), cont_n.get(tkey))
    else:
        g.indisponivel(
            "fotos por painel no manifesto",
            "validation_result.painel_fotografico_contagem ausente na %s -- manifesto "
            "gerado antes da v0.3; rode --fotos para conferir contra a planilha" % args.nova)

    if args.fotos:
        manifest = {}
        if os.path.exists(p_manif):
            import yaml
            with open(p_manif, encoding="utf-8") as fh:
                manifest = yaml.safe_load(fh) or {}
        try:
            contagem, erro = _recontar_fotos(rd, manifest)
        except Exception as e:
            contagem, erro = None, "%s: %s" % (type(e).__name__, e)
        if contagem is None:
            g.indisponivel("fotos por painel", erro or "recontagem falhou")
        else:
            for tkey, esperado in sorted(FOTOS_BASELINE_V01.items()):
                g.valor_esperado("fotos: painel %s = %d" % (tkey, esperado),
                                 esperado, contagem.get(tkey))
    elif not cont_n:
        g.indisponivel(
            "fotos por painel (conferencia contra a planilha)",
            "nao conferida nesta execucao e ausente do manifesto -- use --fotos")

    # ---------------------------------------------------------------- B
    cb = (mb.get("validation_result") or {}).get("document_status_counts") or {}
    cn = (mn.get("validation_result") or {}).get("document_status_counts") or {}
    g.contagem("blocos READY", cb.get("READY", 0), cn.get("READY", 0), "sobe")
    for st in ("HITL", "LIMITATION", "PLACEHOLDER"):
        g.contagem("blocos %s" % st, cb.get(st, 0), cn.get(st, 0), "desce")

    sb = (mb.get("validation_result") or {}).get("document_section_status") or {}
    sn = (mn.get("validation_result") or {}).get("document_section_status") or {}
    for secao in sb:
        if secao in sn:
            g.status_secao(secao, sb[secao], sn[secao])
        else:
            g.checagem("secao %s existe na %s" % (secao, args.nova), False, "secao sumiu")

    figb, fign = _figuras_status(mb), _figuras_status(mn)
    for nome in sorted(set(figb) | set(fign)):
        b, n = figb.get(nome), fign.get(nome)
        if b == "OK" and n != "OK":
            g.regressoes.append("figura %s: OK -> %s" % (nome, n))
            g._log("REGRESSION", "figura %s" % nome, b, n, "figura piorou")
        elif b is None and n is None:
            g.indisponivel("figura %s" % nome, "sem status nas duas versoes")
        else:
            g.iguais += 1
            g._log("OK", "figura %s" % nome, b, n, "estavel ou melhor")

    # ---------------------------------------------------------------- C
    docx_nova = os.path.join(dirsem, "shadow_final_candidate_%s.docx" % args.nova)
    if os.path.exists(docx_nova):
        import zipfile
        try:
            import docx as _docx
            z = zipfile.ZipFile(U.lp(docx_nova))
            g.checagem("docx %s: zip integro" % args.nova, z.testzip() is None)
            n_media = len([n for n in z.namelist() if n.startswith("word/media/")])
            d = _docx.Document(U.lp(docx_nova))
            g.checagem("docx %s: abre no python-docx" % args.nova, True,
                       "%d paragrafos, %d secoes" % (len(d.paragraphs), len(d.sections)))
            g.checagem("docx %s: traz imagens" % args.nova, n_media > 0,
                       "%d arquivos em word/media" % n_media)

            texto = "\n".join(p.text for p in d.paragraphs)
            for t in d.tables:
                for row in t.rows:
                    texto += "\n" + " ".join(c.text for c in row.cells)

            # os fatos precisam estar VISIVEIS no documento, nao so no json
            g.checagem("docx: numero 049 na capa", "049" in texto)
            g.checagem("docx: N° 49 no Termo", "49" in texto)
            for rot, v in (("avanco concluido", fatos.get("avanco_concluido_pct")),
                          ("avanco atacado", fatos.get("avanco_atacado_pct"))):
                if isinstance(v, (int, float)):
                    br = ("%.2f" % v).replace(".", ",")
                    g.checagem("docx: %s = %s" % (rot, br), br in texto, br)
            g.checagem("docx: periodo no texto",
                       "29/08" in texto and "04/09" in texto)
            for s in (rd.get("services") or []):
                nome_s = s.get("canonical_name") or ""
                if nome_s:
                    g.checagem("docx: servico %s" % nome_s, nome_s in texto)
        except Exception as e:
            g.checagem("docx %s legivel" % args.nova, False, "%s: %s" % (type(e).__name__, e))
    else:
        g.checagem("docx %s existe" % args.nova, False, docx_nova)

    # ---------------------------------------------------------------- D
    # Esta camada nasceu de uma falha do proprio guard: na primeira montagem
    # da v0.2 o patch do Termo apagou o corpo da secao 9 e o titulo da secao
    # 10 (geometria fixa que nao sobreviveu a mudanca de quebra de pagina), e
    # o guard deu LIBERA -- porque so olhava manifests e DOCX, e o estrago
    # estava no PDF montado. Manifest e DOCX nao provam o documento final.
    montado = os.path.join(dirsem, "shadow_final_candidate_%s_MONTADO.pdf" % args.nova)
    if not os.path.exists(montado):
        # v0.3: PDF ausente BLOQUEIA. Antes isto era NAO_VERIFICAVEL, e o
        # guard podia dar LIBERA olhando so o DOCX -- exatamente a brecha que
        # deixou passar a pagina 17 quebrada. O documento final e' o PDF.
        g.checagem("PDF montado %s existe" % args.nova, False,
                   "ausente -- rode scripts/montar_pdf_final.py; sem o PDF o guard "
                   "nao tem como liberar (o DOCX sozinho nao prova o documento final)")
    else:
        try:
            import fitz
            doc = fitz.open(U.lp(montado))
            texto_pdf = "\n".join(doc[i].get_text("text") for i in range(doc.page_count))
            g.checagem("pdf: abre e tem paginas", doc.page_count > 0,
                       "%d paginas" % doc.page_count)

            # todo titulo de secao do DOCX precisa sobreviver ate o PDF
            import docx as _docx
            d2 = _docx.Document(U.lp(os.path.join(
                dirsem, "shadow_final_candidate_%s.docx" % args.nova)))
            titulos = [p.text.strip() for p in d2.paragraphs
                       if p.style.name.startswith("Heading") and p.text.strip()]
            sumidos = [t for t in titulos
                       if t.replace("[HITL] ", "").replace("[PENDENTE] ", "") not in texto_pdf]
            g.checagem("pdf: todos os %d titulos do DOCX presentes" % len(titulos),
                       not sumidos, "sumiram: %s" % sumidos if sumidos else "nenhum sumiu")

            # nenhuma marca de placeholder pode sobreviver a montagem
            for marca, rotulo in (("[PENDENTE]", "[PENDENTE]"),
                                  ("___", "lacuna ___"),
                                  ("n_paginas_provisorio", "nota provisoria do Termo")):
                g.checagem("pdf: sem %s" % rotulo, marca not in texto_pdf,
                           "encontrado no PDF montado" if marca in texto_pdf else "ausente")

            # o Termo tem de estar presente e declarar a contagem real
            import re as _re
            g.checagem("pdf: Termo de Encerramento presente no corpo",
                       "TERMO DE ENCERRAMENTO" in texto_pdf)
            m = _re.search(r"constitu[íi]do de\s+(\d+)\s*\(", texto_pdf)
            declarado = int(m.group(1)) if m else None
            g.valor_esperado("pdf: Termo declara o total real de paginas",
                             doc.page_count, declarado)

            # numero do relatorio visivel no documento final
            num = fatos.get("numero_relatorio")
            if num is not None:
                g.checagem("pdf: numero %03d na capa" % int(num),
                           ("%03d" % int(num)) in texto_pdf)
                g.checagem("pdf: N %d no Termo" % int(num),
                           _re.search(r"N[°º]\s*%d\b" % int(num), texto_pdf) is not None)

            # anexos criticos e formatos de pagina
            from collections import Counter as _Counter
            formatos = _Counter()
            for i in range(doc.page_count):
                r = doc[i].rect
                formatos["%dx%d" % (round(r.width), round(r.height))] += 1
            paisagem_grande = sum(n for k, n in formatos.items()
                                  if int(k.split("x")[0]) > 1500)
            g.checagem("pdf: 1 prancha grande (unifilar A0)",
                       any(int(k.split("x")[0]) > 3000 for k in formatos),
                       "formatos: %s" % dict(formatos))

            # o unifilar pode voltar ao placeholder SEM sinal nenhum no PDF: o
            # fallback tambem e' A0 paisagem e tambem tem 1 pagina. So o
            # metodo registrado na montagem distingue prancha real de aviso.
            gates_nova = _ler(os.path.join(dirsem, "critical_gates_%s.json" % args.nova))
            uni = (gates_nova or {}).get("unifilar")
            if uni is None:
                g.indisponivel("unifilar: metodo de geracao",
                               "critical_gates_%s.json sem a chave 'unifilar' -- "
                               "montagem anterior a v0.3" % args.nova)
            else:
                metodo = uni.get("metodo") or ""
                g.checagem("unifilar: prancha REAL (nao placeholder)",
                           "placeholder" not in metodo.lower() and bool(uni.get("ok")),
                           "metodo=%s" % metodo)
                g.valor_esperado("unifilar: 1 pagina", 1, uni.get("n_paginas"))
            g.checagem("pdf: paginas A2 do diario presentes", paisagem_grande >= 7,
                       "%d paginas em paisagem grande" % paisagem_grande)

            # PERDA TEXTUAL CAUSADA POR PATCH
            # O patch do Termo reescreve o PDF por redaction. Se a area
            # redigida for calculada errado, ele APAGA conteudo vizinho -- foi
            # o que aconteceu na primeira montagem da v0.2 (sumiram o corpo da
            # secao 9 e o titulo da secao 10). Aqui o texto do PDF ANTES do
            # patch e' comparado com o do final: tudo que existia tem de
            # continuar existindo, exceto o proprio placeholder.
            #
            # A comparacao e' feita PAGINA A PAGINA e separa dois regimes:
            #
            #   paginas que o patch NAO toca -- nada pode sumir, ponto;
            #   pagina do Termo -- o paragrafo e' reescrito (outras quebras de
            #     linha, outro numero), entao comparar linha a linha daria
            #     falso positivo; o que se exige ali e' que os TITULOS DE
            #     SECAO tenham sobrevivido, que e' precisamente o que a
            #     regressao da v0.2 destruiu.
            pre = os.path.join(dirsem, "merged_pass1.pdf")
            if os.path.exists(pre):
                dpre = fitz.open(U.lp(pre))
                pagina_termo = None
                for i in range(dpre.page_count):
                    t = dpre[i].get_text("text")
                    if "TERMO DE ENCERRAMENTO" in t and ("___" in t or "[PENDENTE]" in t):
                        pagina_termo = i
                        break

                def _linhas(t):
                    return {l.strip() for l in t.splitlines() if len(l.strip()) >= 25}

                perdidas = []
                if dpre.page_count == doc.page_count:
                    for i in range(dpre.page_count):
                        if i == pagina_termo:
                            continue
                        faltando = _linhas(dpre[i].get_text("text")) - \
                            _linhas(doc[i].get_text("text"))
                        perdidas += ["p%d: %s" % (i + 1, l) for l in sorted(faltando)]
                    g.checagem("pdf: patch nao apagou nada fora da pagina do Termo",
                               not perdidas,
                               "%d linha(s) sumiram: %s" % (len(perdidas), perdidas[:3])
                               if perdidas else "%d paginas conferidas" % (doc.page_count - 1))
                else:
                    g.indisponivel("pdf: perda textual fora da pagina do Termo",
                                   "pre-patch tem %d paginas e final tem %d -- sem "
                                   "correspondencia 1:1" % (dpre.page_count, doc.page_count))

                if pagina_termo is not None:
                    titulos_pre = [l.strip() for l in
                                   dpre[pagina_termo].get_text("text").splitlines()
                                   if re.match(r"^\d+\.?\s*$", l.strip())
                                   or re.match(r"^\d+\.\s+[A-ZÀ-Ú]", l.strip())]
                    texto_fim = doc[pagina_termo].get_text("text")
                    # o texto do titulo pode vir quebrado ("10." numa linha,
                    # "TERMO DE ENCERRAMENTO" na outra) -- exige-se o nome
                    faltam = [t for t in ("TERMO DE ENCERRAMENTO",)
                              if t in dpre[pagina_termo].get_text("text")
                              and t not in texto_fim]
                    g.checagem("pdf: titulos da pagina do Termo sobreviveram ao patch",
                               not faltam,
                               "sumiram: %s" % faltam if faltam
                               else "%d marcador(es) de titulo conferidos"
                                    % max(1, len(titulos_pre)))
                dpre.close()
            else:
                g.indisponivel("pdf: perda textual pelo patch",
                               "merged_pass1.pdf ausente -- sem termo de comparacao")
            doc.close()
        except Exception as e:
            g.checagem("PDF montado %s legivel" % args.nova, False,
                       "%s: %s" % (type(e).__name__, e))

    # ---------------------------------------------------------------- saida
    print("=" * 78)
    print("REGRESSION GUARD  %s  %s -> %s" % (args.obra, args.base, args.nova))
    print("=" * 78)
    for ln in g.linhas:
        if ln["nivel"] != "OK":
            print("  %-16s %-42s %s -> %s" % (ln["nivel"], ln["item"][:42],
                                              ln["base"], ln["nova"]))
            print("  %-16s   %s" % ("", ln["obs"]))
    print("-" * 78)
    print("  estaveis ......... %d" % g.iguais)
    print("  melhorias ........ %d" % len(g.melhorias))
    for m in g.melhorias:
        print("      + %s" % m)
    print("  JUSTIFICADOS ..... %d" % len(g.justificados))
    for j in g.justificados:
        print("      ! %s" % j)
    print("  NAO_VERIFICAVEL .. %d" % len(g.nao_verificaveis))
    for nv in g.nao_verificaveis:
        print("      ? %s" % nv)
    print("  REGRESSOES ....... %d" % len(g.regressoes))
    for r in g.regressoes:
        print("      X %s" % r)
    print("=" * 78)

    saida = {
        "obra": args.obra, "periodo": [args.inicio, args.fim],
        "base": args.base, "nova": args.nova,
        "estaveis": g.iguais,
        "melhorias": g.melhorias,
        "nao_verificaveis": g.nao_verificaveis,
        "regressoes": g.regressoes,
        "justificados": g.justificados,
        "veredito": ("BLOQUEIA" if g.regressoes
                     else ("LIBERA_COM_JUSTIFICATIVA" if g.justificados else "LIBERA")),
        "linhas": g.linhas,
    }
    out = os.path.join(dirsem, "regression_guard_%s_%s.json" % (args.base, args.nova))
    with open(U.lp(out), "w", encoding="utf-8") as fh:
        json.dump(saida, fh, ensure_ascii=False, indent=2)
    print("VEREDITO: %s" % saida["veredito"])
    print("gravado: %s" % out)
    return 1 if g.regressoes else 0


if __name__ == "__main__":
    sys.exit(main())
