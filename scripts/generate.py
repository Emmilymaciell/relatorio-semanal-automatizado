# -*- coding: utf-8 -*-
"""
generate.py - COMANDO UNICO do Relatorio Semanal Programa (camada de orquestracao).

    python scripts/generate.py --road RD103 --start 2026-08-29 --end 2026-09-04 \
        [--versao mr1] [--modo HISTORICAL_SHADOW|CURRENT] [--base-versao v0.2] \
        [--skip-pdf] [--pdf-manual] [--resume-from assemble] [--fotos] [--extrair-paginas]

Encadeia, SEM alterar o motor:
    1. candidate  -> scripts/shadow_report.py candidate   (DOCX + manifest + provenance)
    2. pdf        -> scripts/gerador/docx_to_pdf.py       (Word COM isolado; DOCX -> PDF)
    3. assemble   -> scripts/montar_pdf_final.py          (anexos + Termo + validador 15 regras
                                                           + 12 critical gates + PDF MONTADO)
    4. guard      -> scripts/21_regression_guard.py       (so com --base-versao)
    5. scorecard  -> generation_summary_<v>.json + scorecard_<v>.json + SCORECARD_<v>.md

Exit code: 0 = READY_FOR_SHADOW_REVIEW | 1 = BLOCKED | 2 = argumentos/ambiente invalidos
           3 = PDF manual pendente (retomar com --resume-from assemble).

Nunca publica nada. Toda escrita fica dentro do repositorio (prog_util.assert_read_only).
"""
import argparse
import datetime as dt
import glob
import hashlib
import json
import os
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
sys.path.insert(0, os.path.join(HERE, "gerador"))

import prog_util as U  # noqa: E402

ETAPAS = ["candidate", "pdf", "assemble", "guard", "scorecard"]
VEREDITO_READY = "READY_FOR_SHADOW_REVIEW"
VEREDITO_BLOCKED = "BLOCKED"
EXIT_READY, EXIT_BLOCKED, EXIT_ARGS, EXIT_PDF_MANUAL = 0, 1, 2, 3
GATE_PASS = ("CRITICAL_GATE_PASS", "CRITICAL_GATE_PASS_COM_EXCECAO", "PASS")
ETAPA_OK = ("OK", "SKIPPED")


# ----------------------------------------------------------------------------- argumentos
def normalizar_obra(road):
    """Aceita RD103, RD-103, rd_103 -> chave canonica de U.OBRAS."""
    chave = "".join(ch for ch in road.upper() if ch.isalnum())
    if chave not in U.OBRAS:
        raise ValueError("obra desconhecida: %r (validas: %s)" % (road, ", ".join(U.OBRAS)))
    return chave


def parse_args(argv=None):
    ap = argparse.ArgumentParser(
        description="Comando unico: gera, converte, monta e valida o Relatorio Semanal PROG.")
    ap.add_argument("--road", required=True, help="RD101 | RD102A | RD103 | RD104")
    ap.add_argument("--start", required=True, help="AAAA-MM-DD (sabado)")
    ap.add_argument("--end", default=None, help="AAAA-MM-DD (default: start + 6 dias)")
    ap.add_argument("--versao", default="mr1", help="rotulo da rodada (default mr1)")
    ap.add_argument("--modo", default="HISTORICAL_SHADOW", choices=["CURRENT", "HISTORICAL_SHADOW"])
    ap.add_argument("--base-versao", default=None,
                    help="se informada, roda o regression guard contra essa versao")
    ap.add_argument("--skip-pdf", action="store_true",
                    help="nao converte; usa o PDF ja existente no outdir (ou para, se nao houver)")
    ap.add_argument("--pdf-manual", action="store_true",
                    help="para apos o DOCX e imprime a instrucao de conversao manual")
    ap.add_argument("--resume-from", default=None, choices=["candidate", "pdf", "assemble", "guard"],
                    help="pula as etapas anteriores (artefatos precisam existir no outdir)")
    ap.add_argument("--pdf-timeout", type=int, default=900, help="timeout da conversao Word (s)")
    ap.add_argument("--fotos", action="store_true", help="repassa --fotos ao guard (le planilha de 160 MB)")
    ap.add_argument("--extrair-paginas", action="store_true",
                    help="roda scripts/19_extrair_paginas_candidate.py ao final da montagem")
    a = ap.parse_args(argv)
    a.obra = normalizar_obra(a.road)
    a.inicio = dt.date.fromisoformat(a.start)
    a.fim = dt.date.fromisoformat(a.end) if a.end else a.inicio + dt.timedelta(days=6)
    if a.fim < a.inicio:
        ap.error("--end anterior a --start")
    return a


def outdir_de(obra, inicio, fim):
    return os.path.join(U.OUT, "shadow", "%s_%s_%s" % (obra, inicio.isoformat(), fim.isoformat()))


# ----------------------------------------------------------------------------- utilidades
def _agora():
    return dt.datetime.now().astimezone().isoformat(timespec="seconds")


def _ler_json(path):
    if not path or not os.path.exists(path):
        return None
    with open(path, encoding="utf-8") as fh:
        return json.load(fh)


def _gravar(path, texto):
    U.assert_read_only(path)
    with open(path, "w", encoding="utf-8", newline="\n") as fh:
        fh.write(texto)
    return path


def _gravar_json(path, obj):
    return _gravar(path, json.dumps(obj, ensure_ascii=False, indent=2, default=str))


def git_head():
    try:
        r = subprocess.run(["git", "rev-parse", "HEAD"], cwd=U.PROJ, capture_output=True,
                           text=True, timeout=20)
        sujo = subprocess.run(["git", "status", "--porcelain", "--", "scripts", "config"],
                              cwd=U.PROJ, capture_output=True, text=True, timeout=20).stdout.strip()
        return {"head": r.stdout.strip() or None, "motor_config_modificados": bool(sujo)}
    except Exception as e:  # noqa: BLE001
        return {"head": None, "erro": "%s: %s" % (type(e).__name__, e)}


def sha1_configs():
    out = {}
    for p in sorted(glob.glob(os.path.join(U.PROJ, "config", "*.yaml"))):
        with open(p, "rb") as fh:
            out[os.path.basename(p)] = hashlib.sha1(fh.read()).hexdigest()
    return out


def _run(cmd, log_path):
    """Executa cmd, ecoa a saida em tempo real e a anexa ao log. Devolve (rc, ultimas_linhas)."""
    env = dict(os.environ, PYTHONIOENCODING="utf-8", PYTHONUTF8="1", PYTHONUNBUFFERED="1")
    U.assert_read_only(log_path)
    ultimas = []
    with open(log_path, "a", encoding="utf-8") as lg:
        lg.write("\n$ %s\n" % " ".join(cmd))
        p = subprocess.Popen(cmd, cwd=U.PROJ, env=env, stdout=subprocess.PIPE,
                             stderr=subprocess.STDOUT, text=True, encoding="utf-8", errors="replace")
        for linha in p.stdout:
            sys.stdout.write(linha)
            lg.write(linha)
            ultimas.append(linha.rstrip("\n"))
            ultimas = ultimas[-40:]
        p.wait()
        lg.write("[rc=%d]\n" % p.returncode)
    sys.stdout.flush()
    return p.returncode, ultimas


def _py(*args):
    return [sys.executable, "-X", "utf8"] + list(args)


# ----------------------------------------------------------------------------- scorecard (puro)
def decidir_veredito(validador, gates, etapas, guard=None):
    """Regra mecanica de veredito. Entradas sao dicts/listas ja lidos de JSON.

    BLOCKED se: etapa em erro; validador ausente ou com qualquer FAIL; gates ausentes ou
    qualquer gate != PASS (inclui N/A); guard presente com veredito BLOQUEIA/regressoes.
    Devolve (veredito, motivos[])."""
    motivos = []
    for nome, st in (etapas or {}).items():
        if st.get("status") not in ETAPA_OK:
            motivos.append("etapa %s: %s" % (nome, st.get("status")))
    if not validador or not validador.get("resultados"):
        motivos.append("validador ausente (validacao_pdf_final_<v>.json)")
    else:
        for r in validador["resultados"]:
            if r.get("status") != "PASS":
                motivos.append("validador %s %s" % (r.get("id"), r.get("status")))
    if not gates or not gates.get("gates"):
        motivos.append("critical gates ausentes (critical_gates_<v>.json)")
    else:
        for g in gates["gates"]:
            if g.get("status") not in GATE_PASS:
                motivos.append("gate %s (%s) %s" % (g.get("id"), g.get("gate"), g.get("status")))
    if guard is not None:
        if guard.get("veredito") == "BLOQUEIA" or guard.get("regressoes"):
            motivos.append("guard: %s (%d regressao(oes))" % (
                guard.get("veredito"), len(guard.get("regressoes") or [])))
    return (VEREDITO_BLOCKED if motivos else VEREDITO_READY), motivos


def listar_pendencias(manifest, marcadores=None):
    """HITL/LIMITATION/PLACEHOLDER legiveis: secoes nao-READY + figuras fora de OK + marcadores do DOCX."""
    vr = (manifest or {}).get("validation_result") or {}
    itens = []
    for secao, st in (vr.get("document_section_status") or {}).items():
        if st != "READY":
            itens.append({"tipo": st, "onde": secao, "origem": "document_section_status"})
    figs = vr.get("figuras_frente_a") or {}

    def _rec(prefixo, v):
        if isinstance(v, dict):
            for k, x in v.items():
                _rec("%s.%s" % (prefixo, k) if prefixo else k, x)
        elif v not in ("OK", None):
            itens.append({"tipo": "FIGURA_%s" % v, "onde": prefixo, "origem": "figuras_frente_a"})
    _rec("", figs)
    for m in (marcadores or []):
        itens.append({"tipo": "MARCADOR_DOCX", "onde": m, "origem": "docx"})
    return itens


def montar_scorecard(*, obra, periodo, versao, base_versao, validador, gates, guard, manifest,
                     etapas, marcadores=None, n_paginas=None):
    veredito, motivos = decidir_veredito(validador, gates, etapas, guard)
    res = (validador or {}).get("resultados") or []
    gl = (gates or {}).get("gates") or []
    vr = (manifest or {}).get("validation_result") or {}
    pendencias = listar_pendencias(manifest, marcadores)
    pags = n_paginas
    if pags is None:
        pags = (gates or {}).get("n_paginas_final") or (validador or {}).get("n_paginas")
    return {
        "gerado_em": _agora(), "obra": obra, "periodo": periodo, "versao": versao,
        "base_versao": base_versao, "veredito": veredito, "motivos_bloqueio": motivos,
        "n_paginas": pags,
        "validador": {
            "pass": sum(1 for r in res if r.get("status") == "PASS"),
            "fail": sum(1 for r in res if r.get("status") == "FAIL"),
            "outros": sum(1 for r in res if r.get("status") not in ("PASS", "FAIL")),
            "total": len(res),
            "regras": [{"id": r.get("id"), "regra": r.get("regra"), "status": r.get("status"),
                        "detalhe": r.get("detalhe")} for r in res]},
        "gates": {
            "pass": sum(1 for g in gl if g.get("status") in GATE_PASS),
            "total": len(gl),
            "itens": [{"id": g.get("id"), "gate": g.get("gate"), "status": g.get("status"),
                       "evidencia": g.get("evidencia")} for g in gl]},
        "guard": None if guard is None else {
            "veredito": guard.get("veredito"), "base": guard.get("base"), "nova": guard.get("nova"),
            "estaveis": guard.get("estaveis"),
            "melhorias": len(guard.get("melhorias") or []),
            "justificados": len(guard.get("justificados") or []),
            "nao_verificaveis": len(guard.get("nao_verificaveis") or []),
            "regressoes": list(guard.get("regressoes") or [])},
        "blocos_por_status": vr.get("document_status_counts") or {},
        "secoes_por_status": vr.get("document_section_status") or {},
        "pendencias": pendencias,
        "etapas": {k: {"status": v.get("status")} for k, v in (etapas or {}).items()},
    }


def scorecard_md(sc):
    L = []
    L.append("# SCORECARD %s — %s %s..%s" % (sc["versao"], sc["obra"], sc["periodo"]["start"],
                                             sc["periodo"]["end"]))
    L.append("")
    L.append("**Veredito: `%s`**  ·  gerado em %s  ·  páginas: %s" % (
        sc["veredito"], sc["gerado_em"], sc.get("n_paginas")))
    if sc["motivos_bloqueio"]:
        L.append("")
        L.append("Motivos do bloqueio:")
        for m in sc["motivos_bloqueio"]:
            L.append("- %s" % m)
    L.append("")
    L.append("## Etapas")
    for k, v in sc["etapas"].items():
        L.append("- %s: %s" % (k, v["status"]))
    v = sc["validador"]
    L.append("")
    L.append("## Validador (%d PASS / %d FAIL / %d total)" % (v["pass"], v["fail"], v["total"]))
    L.append("")
    L.append("| Regra | Descrição | Status |")
    L.append("|---|---|---|")
    for r in v["regras"]:
        L.append("| %s | %s | %s |" % (r["id"], (r["regra"] or "").replace("|", "/"), r["status"]))
    g = sc["gates"]
    L.append("")
    L.append("## Critical gates (%d/%d PASS)" % (g["pass"], g["total"]))
    L.append("")
    L.append("| Gate | Nome | Status |")
    L.append("|---|---|---|")
    for x in g["itens"]:
        L.append("| %s | %s | %s |" % (x["id"], x["gate"], x["status"]))
    L.append("")
    L.append("## Regression guard")
    gd = sc["guard"]
    if gd is None:
        L.append("Não executado (sem `--base-versao`).")
    else:
        L.append("%s → %s: **%s** · estáveis %s · melhorias %s · justificados %s · "
                 "não verificáveis %s · regressões %d" % (
                     gd["base"], gd["nova"], gd["veredito"], gd["estaveis"], gd["melhorias"],
                     gd["justificados"], gd["nao_verificaveis"], len(gd["regressoes"])))
        for r in gd["regressoes"]:
            L.append("- REGRESSÃO: %s" % r)
    L.append("")
    L.append("## Blocos por status")
    for k, n in sorted(sc["blocos_por_status"].items()):
        L.append("- %s: %s" % (k, n))
    L.append("")
    L.append("## Pendências (HITL / LIMITATION / PLACEHOLDER)")
    if not sc["pendencias"]:
        L.append("Nenhuma.")
    for p in sc["pendencias"]:
        L.append("- [%s] %s" % (p["tipo"], p["onde"]))
    L.append("")
    L.append("> Veredito mecânico. READY_FOR_SHADOW_REVIEW significa apenas \"pronto para revisão "
             "humana\" — nada é publicado automaticamente.")
    return "\n".join(L) + "\n"


# ----------------------------------------------------------------------------- docx markers
def marcadores_docx(docx_path):
    """Paragrafos do DOCX com [HITL]/[PENDENTE] (best effort; nunca falha o pipeline)."""
    try:
        import docx
        d = docx.Document(U.lp(docx_path))
        out = []
        for p in d.paragraphs:
            t = p.text.strip()
            if "[HITL]" in t or "[PENDENTE]" in t:
                out.append(t[:160])
        return out
    except Exception:  # noqa: BLE001
        return []


# ----------------------------------------------------------------------------- pipeline
def instrucao_pdf_manual(docx, pdf, versao, a):
    return ("\n" + "=" * 78 + "\nCONVERSAO MANUAL DE PDF NECESSARIA\n" + "=" * 78 +
            "\n1. Abra no Word (aceite a atualizacao de campos):\n   %s"
            "\n2. Salve/exporte como PDF em:\n   %s"
            "\n3. Retome com:\n   python scripts/generate.py --road %s --start %s --end %s "
            "--versao %s --resume-from assemble%s\n" % (
                docx, pdf, a.obra, a.inicio, a.fim, versao,
                (" --base-versao %s" % a.base_versao) if a.base_versao else "") + "=" * 78 + "\n")


def executar(a):
    obra, versao = a.obra, a.versao
    outdir = outdir_de(obra, a.inicio, a.fim)
    U.assert_read_only(outdir)
    os.makedirs(outdir, exist_ok=True)
    log = os.path.join(outdir, "generate_%s.log" % versao)
    docx = os.path.join(outdir, "shadow_final_candidate_%s.docx" % versao)
    pdf = os.path.join(outdir, "shadow_final_candidate_%s.pdf" % versao)
    montado = os.path.join(outdir, "shadow_final_candidate_%s_MONTADO.pdf" % versao)
    t0 = dt.datetime.now().astimezone()
    resumo = {"comando": " ".join([os.path.basename(sys.argv[0])] + sys.argv[1:]),
              "inicio": t0.isoformat(timespec="seconds"), "git": git_head(),
              "config_sha1": sha1_configs(), "obra": obra, "versao": versao, "modo": a.modo,
              "periodo": {"start": str(a.inicio), "end": str(a.fim)},
              "base_versao": a.base_versao, "python": sys.version.split()[0], "etapas": {}}
    etapas = resumo["etapas"]
    retomar = ETAPAS.index(a.resume_from) if a.resume_from else 0
    if (a.fim - a.inicio).days != 6:
        resumo.setdefault("avisos", []).append("periodo nao tem 7 dias (sabado a sexta)")

    def etapa(nome, status, **extra):
        etapas[nome] = dict(status=status, **extra)

    def finalizar_pdf_manual():
        print(instrucao_pdf_manual(docx, pdf, versao, a))
        resumo["fim"] = _agora()
        resumo["veredito"] = "PDF_MANUAL_PENDENTE"
        _gravar_json(os.path.join(outdir, "generation_summary_%s.json" % versao), resumo)
        return EXIT_PDF_MANUAL

    # 1. candidate
    if retomar > ETAPAS.index("candidate"):
        etapa("candidate", "SKIPPED", motivo="--resume-from")
    else:
        t = dt.datetime.now()
        rc, _ = _run(_py(os.path.join("scripts", "shadow_report.py"), "candidate", "--obra", obra,
                         "--inicio", str(a.inicio), "--fim", str(a.fim), "--versao", versao,
                         "--modo", a.modo), log)
        etapa("candidate", "OK" if rc == 0 else "ERRO", rc=rc,
              segundos=round((dt.datetime.now() - t).total_seconds(), 1))
    if etapas["candidate"]["status"] not in ETAPA_OK:
        return _concluir(a, outdir, resumo, docx, montado, None)
    if not os.path.exists(docx):
        etapa("candidate", "ERRO", motivo="DOCX nao gerado: %s" % docx)
        return _concluir(a, outdir, resumo, docx, montado, None)

    # 2. pdf
    if retomar > ETAPAS.index("pdf"):
        etapa("pdf", "SKIPPED", motivo="--resume-from")
    elif a.pdf_manual:
        etapa("pdf", "PENDENTE_MANUAL")
        return finalizar_pdf_manual()
    elif a.skip_pdf:
        if os.path.exists(pdf):
            etapa("pdf", "SKIPPED", motivo="--skip-pdf, usando PDF existente")
        else:
            etapa("pdf", "PENDENTE_MANUAL", motivo="--skip-pdf sem PDF existente")
            return finalizar_pdf_manual()
    else:
        import docx_to_pdf  # scripts/gerador
        U.assert_read_only(pdf)
        print("[pdf] convertendo via Word COM (~45 s)...")
        r = docx_to_pdf.converter(docx, pdf, timeout=a.pdf_timeout)
        etapa("pdf", "OK" if r.get("status") == "OK" else "ERRO",
              conversao=r.get("status"), segundos=r.get("segundos"),
              n_paginas=(r.get("validacao_pdf") or {}).get("n_paginas"), erro=(r.get("erro") or "")[:500])
        if r.get("status") != "OK":
            print("[pdf] conversao falhou: %s %s" % (r.get("status"), r.get("erro") or r.get("motivo") or ""))
            etapas["pdf"]["status"] = "PENDENTE_MANUAL"
            return finalizar_pdf_manual()
    if not os.path.exists(pdf):
        etapa("pdf", "PENDENTE_MANUAL", motivo="PDF ausente: %s" % pdf)
        return finalizar_pdf_manual()

    # 3. assemble
    if retomar > ETAPAS.index("assemble"):
        etapa("assemble", "SKIPPED", motivo="--resume-from")
    else:
        t = dt.datetime.now()
        rc, _ = _run(_py(os.path.join("scripts", "montar_pdf_final.py"), "--obra", obra,
                         "--inicio", str(a.inicio), "--fim", str(a.fim), "--versao", versao), log)
        etapa("assemble", "OK" if rc == 0 else "ERRO", rc=rc,
              segundos=round((dt.datetime.now() - t).total_seconds(), 1))
        if rc == 0 and a.extrair_paginas:
            rc2, _ = _run(_py(os.path.join("scripts", "19_extrair_paginas_candidate.py"),
                              "--versao", versao), log)
            etapa("extrair_paginas", "OK" if rc2 == 0 else "ERRO", rc=rc2)

    # 4. guard
    if not a.base_versao:
        etapa("guard", "SKIPPED", motivo="sem --base-versao")
    elif etapas["assemble"]["status"] not in ETAPA_OK:
        etapa("guard", "SKIPPED", motivo="montagem falhou")
    else:
        cmd = _py(os.path.join("scripts", "21_regression_guard.py"), "--obra", obra,
                  "--inicio", str(a.inicio), "--fim", str(a.fim), "--base", a.base_versao,
                  "--nova", versao)
        if a.fotos:
            cmd.append("--fotos")
        rc, _ = _run(cmd, log)
        # rc 1 = regressao detectada (resultado valido, entra no veredito); rc>=2 = erro de execucao
        etapa("guard", "OK" if rc in (0, 1) else "ERRO", rc=rc)

    return _concluir(a, outdir, resumo, docx, montado, True)


def _concluir(a, outdir, resumo, docx, montado, _tentar_scorecard):
    obra, versao = a.obra, a.versao
    P = lambda n: os.path.join(outdir, n)  # noqa: E731
    etapas = resumo["etapas"]
    validador = _ler_json(P("validacao_pdf_final_%s.json" % versao)) if etapas.get("assemble", {}).get("status") in ETAPA_OK else None
    gates = _ler_json(P("critical_gates_%s.json" % versao)) if validador is not None else None
    guard = None
    if a.base_versao and etapas.get("guard", {}).get("status") == "OK":
        guard = _ler_json(P("regression_guard_%s_%s.json" % (a.base_versao, versao)))
        if guard is None:
            etapas["guard"]["status"] = "ERRO"
    manifest = _ler_json(P("candidate_manifest_%s.json" % versao))
    marc = marcadores_docx(docx) if os.path.exists(docx) else []
    etapas_sc = {k: v for k, v in etapas.items()}
    sc = montar_scorecard(obra=obra, periodo={"start": str(a.inicio), "end": str(a.fim)},
                          versao=versao, base_versao=a.base_versao, validador=validador, gates=gates,
                          guard=guard, manifest=manifest, etapas=etapas_sc, marcadores=marc)
    p_sc = _gravar_json(P("scorecard_%s.json" % versao), sc)
    p_md = _gravar(P("SCORECARD_%s.md" % versao), scorecard_md(sc))
    etapas["scorecard"] = {"status": "OK"}

    vr = (manifest or {}).get("validation_result") or {}
    contagem = vr.get("document_status_counts") or {}
    resumo.update({
        "fim": _agora(), "veredito": sc["veredito"], "motivos_bloqueio": sc["motivos_bloqueio"],
        "caminhos": {
            "docx": docx, "pdf_conversao": os.path.join(outdir, "shadow_final_candidate_%s.pdf" % versao),
            "pdf_montado": montado if os.path.exists(montado) else None,
            "manifest": P("candidate_manifest_%s.json" % versao), "manifest_yaml": P("manifest.yaml"),
            "provenance": P("report_provenance.json"), "report_data": P("report_data.json"),
            "validacao": P("validacao_pdf_final_%s.json" % versao),
            "gates": P("critical_gates_%s.json" % versao),
            "guard": P("regression_guard_%s_%s.json" % (a.base_versao, versao)) if a.base_versao else None,
            "scorecard_json": p_sc, "scorecard_md": p_md, "log": P("generate_%s.log" % versao)},
        "contagem": {k: contagem.get(k, 0) for k in ("READY", "HITL", "LIMITATION", "PLACEHOLDER")},
        "pendencias": [("[%s] %s" % (p["tipo"], p["onde"])) for p in sc["pendencias"]],
        "n_paginas": sc["n_paginas"]})
    _gravar_json(P("generation_summary_%s.json" % versao), resumo)

    print("\n" + "=" * 78)
    print("VEREDITO: %s   (%s %s..%s, versao %s)" % (sc["veredito"], obra, a.inicio, a.fim, versao))
    print("validador: %d PASS / %d FAIL  |  gates: %d/%d  |  paginas: %s  |  guard: %s" % (
        sc["validador"]["pass"], sc["validador"]["fail"], sc["gates"]["pass"], sc["gates"]["total"],
        sc["n_paginas"], (sc["guard"] or {}).get("veredito", "nao executado")))
    for m in sc["motivos_bloqueio"]:
        print("  BLOQUEIO: %s" % m)
    print("scorecard: %s" % p_md)
    print("resumo:    %s" % P("generation_summary_%s.json" % versao))
    print("Nada foi publicado. Revisao humana necessaria.")
    print("=" * 78)
    return EXIT_READY if sc["veredito"] == VEREDITO_READY else EXIT_BLOCKED


def main(argv=None):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:  # noqa: BLE001
        pass
    try:
        a = parse_args(argv)
    except ValueError as e:
        print("ERRO: %s" % e, file=sys.stderr)
        return EXIT_ARGS
    return executar(a)


if __name__ == "__main__":
    sys.exit(main())
