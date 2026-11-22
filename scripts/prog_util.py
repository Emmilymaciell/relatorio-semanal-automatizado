# -*- coding: utf-8 -*-
r"""
prog_util.py - utilitarios comuns do projeto relatorio-semanal-automatizado.

REGRA CRITICA: o SharePoint (C:\caminho\exemplo\PastaSincronizada\) e READ-ONLY.
Nenhuma funcao deste modulo escreve nele. `assert_read_only()` protege contra acidentes.

Problema tratado aqui: os caminhos do SharePoint estouram MAX_PATH (260 chars) no Windows,
o que faz os.stat/open falharem com WinError 3. Solucao: prefixo \\?\ (extended-length path)
aplicado por `lp()`. TODO acesso a arquivo do SharePoint deve passar por `lp()`.
"""
import os
import re
import shutil
import hashlib
import unicodedata
from datetime import datetime, timezone

# Raiz do SharePoint sincronizado: sobrescrevivel por PROG_SHAREPOINT (outra maquina/usuario).
SHAREPOINT = os.environ.get("PROG_SHAREPOINT") or os.path.join(
    os.path.expanduser("~"), "PastaSincronizada")
BASE_OBRAS = os.path.join(SHAREPOINT, r"Programa - 3. Desenvolvimento", "G2 -  Engenharia")
# Raiz do repositorio = pasta acima de scripts/ (vale para clone e git worktree).
PROJ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
# tmp/ e' cache de copias do SharePoint (~363 MB); PROG_TMP permite compartilha-lo
# entre worktrees em vez de recopiar.
TMP = os.environ.get("PROG_TMP") or os.path.join(PROJ, "tmp")
OUT = os.path.join(PROJ, "output")
OBRAS = ["RD101", "RD102A", "RD103", "RD104"]
RODOVIA_LABEL = {"RD101": "RD-101", "RD102A": "RD-102A", "RD103": "RD-103", "RD104": "RD-104"}

_PREFIX = "\\\\?\\"      # literal: \\?\
_UNC = "\\\\?\\UNC\\"    # literal: \\?\UNC\


def lp(path):
    r"""Long path: prefixa \\?\ para driblar MAX_PATH no Windows."""
    p = os.path.abspath(path)
    if p.startswith(_PREFIX):
        return p
    if p.startswith("\\\\"):
        return _UNC + p[2:]
    return _PREFIX + p


def unlp(path):
    return path[len(_PREFIX):] if path.startswith(_PREFIX) else path


def assert_read_only(path):
    """Levanta erro se `path` estiver dentro do SharePoint (protecao contra escrita)."""
    p = os.path.abspath(unlp(path)).lower()
    if p.startswith(SHAREPOINT.lower()):
        raise PermissionError("BLOQUEADO: escrita em area READ-ONLY do SharePoint: " + path)
    return path


def listdir(path):
    return sorted(os.listdir(lp(path)))


def isdir(path):
    return os.path.isdir(lp(path))


def exists(path):
    return os.path.exists(lp(path))


def walk(path):
    """os.walk seguro para caminhos longos; devolve dirpath sem o prefixo."""
    for dirpath, dirnames, filenames in os.walk(lp(path)):
        dirnames[:] = sorted(dirnames)
        yield unlp(dirpath), dirnames, sorted(filenames)


def stat_info(path):
    st = os.stat(lp(path))
    return {
        "tamanho_bytes": st.st_size,
        "modificado": datetime.fromtimestamp(st.st_mtime, timezone.utc)
                              .astimezone().isoformat(timespec="seconds"),
    }


def copiar_para_tmp(origem, subdir="", nome=None):
    """Copia um arquivo do SharePoint para tmp/ do projeto. Nunca escreve na origem."""
    destino_dir = os.path.join(TMP, subdir)
    assert_read_only(destino_dir)
    os.makedirs(destino_dir, exist_ok=True)
    destino = os.path.join(destino_dir, nome or os.path.basename(origem))
    # O destino tem nome fixo por subdir (ex.: tmp/avanco/prog_avanco.xlsx) e e'
    # reaproveitado entre obras e semanas. So tamanho nao basta: dois arquivos
    # diferentes do mesmo tamanho seriam confundidos e o extrator leria a fonte de
    # outra obra/semana em silencio. A copia preserva o mtime (copystat) e o cache so
    # vale se tamanho E mtime baterem.
    st = os.stat(lp(origem))
    valido = False
    if os.path.exists(lp(destino)):
        sd = os.stat(lp(destino))
        valido = sd.st_size == st.st_size and int(sd.st_mtime) == int(st.st_mtime)
    if not valido:
        with open(lp(origem), "rb") as fi, open(lp(destino), "wb") as fo:
            shutil.copyfileobj(fi, fo, 1024 * 1024)
        os.utime(lp(destino), (st.st_atime, st.st_mtime))
    return destino


def sha1(path, limite=None):
    h = hashlib.sha1()
    with open(lp(path), "rb") as fh:
        lido = 0
        while True:
            b = fh.read(1024 * 1024)
            if not b:
                break
            h.update(b)
            lido += len(b)
            if limite and lido >= limite:
                break
    return h.hexdigest()


def norm(s):
    """Normaliza texto para comparacao: sem acento, minusculo, espacos colapsados."""
    s = unicodedata.normalize("NFKD", str(s))
    s = "".join(c for c in s if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", s).strip().lower()


def slug(s):
    return re.sub(r"[^a-z0-9]+", "_", norm(s)).strip("_")


def pasta_semanal_raiz(obra):
    """Pasta '4. Relatorio Semanal (de) Obras' da obra."""
    gto = os.path.join(BASE_OBRAS, obra, "3_Gestao Tecnica de Obra")
    for d in listdir(gto):
        if d.startswith("4. Relat"):
            return os.path.join(gto, d)
    return None


def pasta_relatorios(obra):
    """Pasta '5. Relatorio Semanal de Acompanhamento de Obras' da obra."""
    raiz = pasta_semanal_raiz(obra)
    if not raiz:
        return None
    for s in listdir(raiz):
        if s.startswith("5. Relat"):
            return os.path.join(raiz, s)
    return None
