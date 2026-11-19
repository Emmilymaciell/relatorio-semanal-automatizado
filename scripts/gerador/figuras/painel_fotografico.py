# -*- coding: utf-8 -*-
r"""
scripts/gerador/figuras/painel_fotografico.py - O ITEM CENTRAL da Fase 3.3:
monta o Relatorio Fotografico (secao 5.2.x) como uma imagem composta por
servico, a partir das fotos JA SELECIONADAS pela engenharia em
`Arquivos Auxiliares/RELATORIO FOTOGRAFICO <n> - <obra>.xlsx`.

## Como as fotos estao armazenadas no arquivo (VERIFICADO nesta sessao)

Nao sao objetos `xl/drawings/drawingN.xml` ancorados (esses so tem o logo do
cabecalho e formas decorativas). As ~60 fotos usam o recurso do Excel 365
"Inserir imagem > Na celula" (`vm="N"` no atributo da celula `<c>`), que vira
uma cadeia de indirecao (rich value):

    celula <c vm="M">           (M = indice 1-based em <valueMetadata>)
      -> xl/metadata.xml <valueMetadata><bk><rc t="1" v="K"/>    (K 0-based)
      -> xl/richData/rdrichvalue.xml <rv> numero K -> <v>L</v>   (LocalImageIdentifier)
      -> xl/richData/richValueRel.xml, L-esimo <rel r:id="rIdX"/>
      -> xl/richData/_rels/richValueRel.xml.rels rIdX -> xl/media/imageNN.<ext>

Isso bate com `scripts/gerador/extratores/fotos.py` (que ja conta `vm="N"`
para decidir subsecao 5.2.x, mas nao extrai a imagem em si).

## Legenda -- achado que MUDA a leitura da spec

A spec (Sec.2.3) supunha que a legenda (Estaca+Data / Coordenada / Descricao)
so existiria como overlay IMPRESSO NO PIXEL da foto (exigindo OCR, fora de
escopo). **Isso nao e verdade para este arquivo**: a mesma celula tambem tem,
na planilha, um bloco de 6 linhas de TEXTO ESTRUTURADO logo acima de cada
foto (verificado em TERRAPLENAGEM/PAVIMENTACAO/DRENAGEM/OAE, RD103/49):

    linha-5: "Estaca:" <estaca>      "Data:" <serial>
    linha-4: "Coordenada:" <utm>
    linha-3: "Sensivel:" Sim/Nao      (hidden=1)
    linha-2: "Indice:" <n>            (hidden=1)
    linha-1: "Descricao:" <texto>
    linha 0: [FOTO]

As 3 linhas VISIVEIS (Estaca+Data / Coordenada / Descricao) sao exatamente as
3 linhas do `GOLDEN_VISUAL_SPEC.md` Sec.7 -- e ja transcritas para texto pela
engenharia, sem precisar de OCR. A leitura aqui e por TEXTO DE CELULA (nao
imagem), localizado pelo ROTULO ("Estaca:", "Data:", ...) numa janela de busca
acima/a direita de cada foto -- nunca por letra de coluna fixa, entao
sobrevive a pequenas variacoes de layout entre abas/semanas.

Quando o rotulo ou o valor correspondente nao e encontrado para uma foto, a
legenda dessa foto fica em branco (moldura com borda, sem texto) e o servico
inteiro e marcado `photo_selection_status: HITL` -- nunca se inventa estaca,
data ou coordenada (regra da spec, Sec.0 e Sec.2.3).

## Contrato

`montar(report_data, manifest, out_dir)` devolve um dict
{chave_servico: FiguraResult}, uma entrada por aba visivel-com-foto do
workbook (a mesma regra de `abas_com_foto`/T-17 de
`extratores/fotos.py` e `config/services_taxonomy.yaml`).
"""
import os
import re
import zipfile
import datetime as dt
import xml.etree.ElementTree as ET
from typing import Dict, List, Optional, Tuple

from openpyxl.utils import column_index_from_string as _cix
from openpyxl.utils import get_column_letter as _col_letter
from PIL import Image, ImageDraw, ImageOps
import io

from . import (FiguraResult, mm_to_px, get_font, wrap_text, salvar_png,
               desenhar_borda, text_width, COR_BORDA, COR_BRANCO, COR_PRETO,
               U, assert_nao_e_golden, C)
from ..photo_panel import PhotoPanelComposer, PANEL_WIDTH_MM

NS_MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
NS_REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
NS_RVREL = "{http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel}"
NS_RD = "{http://schemas.microsoft.com/office/spreadsheetml/2017/richdata}"
NS_XLRD = "{http://schemas.microsoft.com/office/spreadsheetml/2017/richdata}"
NS_PKGREL = "{http://schemas.openxmlformats.org/package/2006/relationships}"

_EXCEL_EPOCH = dt.date(1899, 12, 30)

_ALTURA_MAX_LEGENDA_LINHAS = 3
FONTE_LEGENDA_PT = 7.0
FONTE_LEGENDA_LINHA_MM = 3.6
PAD_CELULA_MM = 1.6
GAP_ENTRE_FOTOS_MM = 1.0


def _ref_to_rowcol(ref: str) -> Tuple[int, int]:
    m = re.match(r"([A-Z]+)(\d+)", ref)
    return int(m.group(2)), _cix(m.group(1))


class _WorkbookIndex:
    """Le uma vez a cadeia inteira (sheets, shared strings, rich value chain)
    de um RELATORIO FOTOGRAFICO <n>.xlsx aberto como zip."""

    def __init__(self, caminho_local: str):
        self.zf = zipfile.ZipFile(caminho_local)
        self._sst = self._ler_sst()
        self._sheets = self._ler_sheets()          # nome -> (path, state)
        self._vm_para_media = self._montar_cadeia_rich_value()

    def _ler_sst(self) -> List[str]:
        if "xl/sharedStrings.xml" not in self.zf.namelist():
            return []
        root = ET.fromstring(self.zf.read("xl/sharedStrings.xml"))
        out = []
        for si in root.findall(NS_MAIN + "si"):
            texts = si.findall(".//" + NS_MAIN + "t")
            out.append("".join(t.text or "" for t in texts))
        return out

    def _ler_sheets(self) -> Dict[str, Tuple[str, str]]:
        wb = ET.fromstring(self.zf.read("xl/workbook.xml"))
        rels = ET.fromstring(self.zf.read("xl/_rels/workbook.xml.rels"))
        relmap = {r.get("Id"): r.get("Target") for r in rels}
        out = {}
        for s in wb.find(NS_MAIN + "sheets"):
            nome = s.get("name")
            rid = s.get(NS_REL + "id")
            estado = s.get("state", "visible")
            target = relmap.get(rid, "")
            parte = "xl/" + target if not target.startswith("/") else target.lstrip("/")
            out[nome] = (parte, estado)
        return out

    def _montar_cadeia_rich_value(self) -> Dict[int, str]:
        """vm (1-based, INDICE GLOBAL do workbook) -> caminho no zip da midia."""
        if "xl/metadata.xml" not in self.zf.namelist():
            return {}
        md = ET.fromstring(self.zf.read("xl/metadata.xml"))
        value_meta = md.find(NS_MAIN + "valueMetadata")
        if value_meta is None:
            return {}
        # valueMetadata[i] (0-based) -> K (indice em rdrichvalue), so para o
        # metadataType que for XLRICHVALUE (na pratica, o unico tipo usado
        # aqui); vm da celula = i+1.
        k_por_indice = []
        for bk in value_meta.findall(NS_MAIN + "bk"):
            rc = bk.find(NS_MAIN + "rc")
            k_por_indice.append(int(rc.get("v")) if rc is not None else None)

        if "xl/richData/rdrichvalue.xml" not in self.zf.namelist():
            return {}
        rv_root = ET.fromstring(self.zf.read("xl/richData/rdrichvalue.xml"))
        l_por_k = []
        for rv in rv_root.findall(NS_RD + "rv"):
            vs = rv.findall(NS_RD + "v")
            l_por_k.append(int(vs[0].text) if vs else None)

        if "xl/richData/richValueRel.xml" not in self.zf.namelist():
            return {}
        rel_root = ET.fromstring(self.zf.read("xl/richData/richValueRel.xml"))
        rid_por_l = [r.get(NS_REL + "id") for r in rel_root.findall(NS_RVREL + "rel")]

        rels_root = ET.fromstring(self.zf.read("xl/richData/_rels/richValueRel.xml.rels"))
        media_por_rid = {}
        for r in rels_root.findall(NS_PKGREL + "Relationship"):
            target = r.get("Target")
            parte = "xl/" + target[3:] if target.startswith("../") else target
            media_por_rid[r.get("Id")] = parte

        vm_para_media = {}
        for idx0, k in enumerate(k_por_indice):
            vm = idx0 + 1
            if k is None or k >= len(l_por_k):
                continue
            l = l_por_k[k]
            if l is None or l >= len(rid_por_l):
                continue
            rid = rid_por_l[l]
            parte = media_por_rid.get(rid)
            if parte:
                vm_para_media[vm] = parte
        return vm_para_media

    def abas_disponiveis(self) -> Dict[str, Tuple[str, str, int]]:
        """nome -> (path, state, n_vm) para todas as abas do workbook."""
        out = {}
        for nome, (path, estado) in self._sheets.items():
            try:
                dados = self.zf.read(path)
            except KeyError:
                out[nome] = (path, estado, 0)
                continue
            n_vm = len(re.findall(rb' vm="\d+"', dados))
            out[nome] = (path, estado, n_vm)
        return out

    def imagem_bytes(self, vm: int) -> Optional[bytes]:
        parte = self._vm_para_media.get(vm)
        if not parte:
            return None
        try:
            return self.zf.read(parte)
        except KeyError:
            return None

    def _resolver_valor_celula(self, c_elem) -> Optional[str]:
        t = c_elem.get("t")
        v_elem = c_elem.find(NS_MAIN + "v")
        is_elem = c_elem.find(NS_MAIN + "is")
        if t == "s" and v_elem is not None:
            try:
                return self._sst[int(v_elem.text)]
            except (ValueError, IndexError):
                return None
        if is_elem is not None:
            tt = is_elem.find(NS_MAIN + "t")
            return tt.text if tt is not None else None
        if v_elem is not None:
            return v_elem.text
        return None

    def ler_celulas_aba(self, path: str):
        """Devolve (grid, vm_cells): grid[(row,col)] = (texto_ou_none, tipo);
        vm_cells = lista [(row,col,vm)] na ordem do documento."""
        ws = ET.fromstring(self.zf.read(path))
        sheetdata = ws.find(NS_MAIN + "sheetData")
        grid = {}
        vm_cells = []
        if sheetdata is None:
            return grid, vm_cells
        for row_el in sheetdata.findall(NS_MAIN + "row"):
            for c_el in row_el.findall(NS_MAIN + "c"):
                ref = c_el.get("r")
                if not ref:
                    continue
                r, col = _ref_to_rowcol(ref)
                vm = c_el.get("vm")
                if vm:
                    vm_cells.append((r, col, int(vm)))
                valor = self._resolver_valor_celula(c_el)
                tipo = c_el.get("t")
                grid[(r, col)] = (valor, tipo)
        vm_cells.sort(key=lambda t: (t[0], t[1]))
        return grid, vm_cells


def _norm_label(txt) -> str:
    if txt is None:
        return ""
    return U.norm(str(txt)).rstrip(":").strip()


def _excel_serial_para_data(v) -> Optional[str]:
    try:
        n = float(v)
    except (TypeError, ValueError):
        return None
    try:
        d = _EXCEL_EPOCH + dt.timedelta(days=n)
        return d.strftime("%d/%m/%Y")
    except OverflowError:
        return None


def _buscar_valor_na_linha(grid, linha, col_inicio, col_fim):
    for c in range(col_inicio, col_fim + 1):
        v, _t = grid.get((linha, c), (None, None))
        if v not in (None, ""):
            return v, c
    return None, None


def _extrair_legenda(grid, row_foto: int, col_foto: int,
                     largura_bloco_cols: int = 12,
                     janela_linhas: int = 8) -> Dict[str, Optional[str]]:
    """Busca, ACIMA da foto (mesmo bloco de colunas), celulas cujo texto
    normalizado bata com um rotulo conhecido ('Estaca:', 'Data:',
    'Coordenada:', 'Descricao:') e le o primeiro valor nao vazio a direita,
    dentro do mesmo bloco de colunas. Generico -- nao assume offset fixo de
    linha/coluna (o layout varia um pouco entre abas)."""
    col_fim = col_foto + largura_bloco_cols - 1
    achados = {"estaca": None, "data": None, "coordenada": None, "descricao": None}
    for r in range(row_foto - 1, max(1, row_foto - janela_linhas) - 1, -1):
        for c in range(col_foto, col_fim + 1):
            v, _t = grid.get((r, c), (None, None))
            rotulo = _norm_label(v)
            if rotulo in achados and achados[rotulo] is None:
                valor, _vc = _buscar_valor_na_linha(grid, r, c + 1, col_fim)
                if rotulo == "data" and valor is not None:
                    valor = _excel_serial_para_data(valor) or str(valor)
                achados[rotulo] = valor
        if all(v is not None for v in achados.values()):
            break
    return achados


def _aba_para_taxonomia() -> Dict[str, str]:
    tax = C.services_taxonomy()
    mapa = tax.get("abas_relatorio_fotografico", {})
    out = {}
    for aba_nome, alvo in mapa.items():
        if isinstance(alvo, str):
            out[U.norm(aba_nome).rstrip(".")] = alvo
    return out


def _nome_canonico(taxonomy_key: str) -> str:
    tax = C.services_taxonomy()
    for grupo in ("servicos", "meio_ambiente_e_sst"):
        bloco = tax.get(grupo, {})
        if taxonomy_key in bloco:
            return bloco[taxonomy_key].get("canonical_name", taxonomy_key)
    return taxonomy_key


def _carregar_imagem(img_bytes: bytes) -> Optional[Image.Image]:
    try:
        im = Image.open(io.BytesIO(img_bytes))
        im = ImageOps.exif_transpose(im)
        return im.convert("RGB")
    except Exception:
        return None


def _slug_arquivo(taxonomy_key: str) -> str:
    return "fotografico_%s" % U.slug(taxonomy_key)


def _desenhar_painel(servico_titulo: str, periodo_txt: str, fotos: List[Dict],
                     n_colunas: int, coluna_mm: float, panel_width_mm: float) -> Image.Image:
    fonte_titulo = get_font(11, bold=True)
    fonte_legenda = get_font(FONTE_LEGENDA_PT)
    fonte_legenda_lbl = get_font(FONTE_LEGENDA_PT, bold=True)

    coluna_px = mm_to_px(coluna_mm)
    largura_px = mm_to_px(panel_width_mm)
    gap_px = mm_to_px(GAP_ENTRE_FOTOS_MM)
    pad_px = mm_to_px(PAD_CELULA_MM)
    linha_legenda_px = mm_to_px(FONTE_LEGENDA_LINHA_MM)
    legenda_altura_px = _ALTURA_MAX_LEGENDA_LINHAS * linha_legenda_px + 2 * pad_px
    foto_area_px = coluna_px - 2 * pad_px

    titulo_altura_px = mm_to_px(9.0)
    n = len(fotos)
    n_linhas = -(-n // n_colunas) if n else 1

    # altura de cada linha do grid = legenda + foto (a foto e redimensionada
    # para caber na largura da coluna, altura livre 50-57mm conforme a spec,
    # usando um valor fixo por painel para manter o grid alinhado)
    altura_foto_mm = 54.0
    altura_foto_px = mm_to_px(altura_foto_mm)
    linha_grid_px = legenda_altura_px + altura_foto_px

    altura_total_px = titulo_altura_px + n_linhas * linha_grid_px
    img = Image.new("RGB", (largura_px, altura_total_px), COR_BRANCO)
    draw = ImageDraw.Draw(img)

    draw.text((mm_to_px(1.5), mm_to_px(1.5)), "%s — %s" % (servico_titulo, periodo_txt),
              font=fonte_titulo, fill=COR_PRETO)
    draw.line([(0, titulo_altura_px), (largura_px, titulo_altura_px)], fill=COR_BORDA, width=1)

    for idx in range(n_linhas * n_colunas):
        linha_i, col_i = divmod(idx, n_colunas)
        x0 = col_i * coluna_px
        y0 = titulo_altura_px + linha_i * linha_grid_px
        x1 = x0 + coluna_px
        y1 = y0 + linha_grid_px
        desenhar_borda(draw, (x0, y0, x1 - 1, y1 - 1), 1, COR_BORDA)

        if idx >= n:
            continue  # celula vazia da ultima linha -- so a borda, sem foto

        foto = fotos[idx]
        # legenda (3 linhas: Estaca+Data / Coordenada / Descricao)
        ty = y0 + pad_px
        linha1 = ""
        if foto.get("estaca") or foto.get("data"):
            partes = []
            if foto.get("estaca"):
                partes.append("Estaca: %s" % foto["estaca"])
            if foto.get("data"):
                partes.append("Data: %s" % foto["data"])
            linha1 = "   ".join(partes)
        linha2 = ("Coordenada: %s" % foto["coordenada"]) if foto.get("coordenada") else ""
        linha3 = foto.get("descricao") or ""

        for texto in (linha1, linha2, linha3):
            if texto:
                ls = wrap_text(draw, texto, fonte_legenda, foto_area_px)
                draw.text((x0 + pad_px, ty), ls[0], font=fonte_legenda, fill=COR_PRETO)
            ty += linha_legenda_px

        # foto (escala para caber na coluna, centralizada; letterbox se
        # preciso para nao distorcer a proporcao original)
        fy0 = y0 + legenda_altura_px
        im = foto.get("_pil")
        if im is not None:
            escala = min(foto_area_px / im.width, altura_foto_px / im.height)
            nw, nh = max(1, int(im.width * escala)), max(1, int(im.height * escala))
            im_r = im.resize((nw, nh))
            px = x0 + pad_px + (foto_area_px - nw) // 2
            py = fy0 + (altura_foto_px - nh) // 2
            img.paste(im_r, (int(px), int(py)))
        else:
            draw.rectangle([x0 + pad_px, fy0, x1 - pad_px, fy0 + altura_foto_px],
                          outline=COR_BORDA)
    return img


def montar(report_data: dict, manifest: dict, out_dir: str) -> Dict[str, FiguraResult]:
    caminho = (((manifest or {}).get("sources") or {}).get("photo_workbook") or {}).get("path")
    resultados: Dict[str, FiguraResult] = {}

    if not caminho:
        return {"_geral": FiguraResult(
            None, PANEL_WIDTH_MM, 0.0, "LIMITATION",
            "manifest.sources.photo_workbook.path ausente -- caminho nao resolvido",
            {"source_file": None})}

    assert_nao_e_golden(caminho)
    if not U.exists(caminho):
        return {"_geral": FiguraResult(
            None, PANEL_WIDTH_MM, 0.0, "LIMITATION",
            "arquivo nao encontrado: %s" % caminho, {"source_file": caminho})}

    destino = "relfoto%s" % os.path.splitext(caminho)[1]
    local = U.copiar_para_tmp(caminho, "fotografico", destino)
    idx = _WorkbookIndex(local)

    abas = idx.abas_disponiveis()
    aba_para_taxonomia = _aba_para_taxonomia()

    periodo = (report_data or {}).get("period", {})
    periodo_txt = "%s a %s" % (periodo.get("start", ""), periodo.get("end", ""))

    servicos_relatorio = {s["key"] for s in (report_data or {}).get("services", [])}

    composer = PhotoPanelComposer()

    candidatas = []
    for aba_nome, (path, estado, n_vm) in abas.items():
        chave_norm = U.norm(aba_nome).rstrip(".")
        taxonomy_key = aba_para_taxonomia.get(chave_norm)
        if not taxonomy_key:
            continue
        if estado != "visible" or n_vm <= 0:
            continue
        candidatas.append((aba_nome, path, taxonomy_key, n_vm))

    for aba_nome, path, taxonomy_key, n_vm in candidatas:
        grid, vm_cells = idx.ler_celulas_aba(path)

        fotos = []
        n_legenda_incompleta = 0
        for (row, col, vm) in vm_cells:
            img_bytes = idx.imagem_bytes(vm)
            pil_img = _carregar_imagem(img_bytes) if img_bytes else None
            legenda = _extrair_legenda(grid, row, col)
            if not (legenda["estaca"] and legenda["data"] and legenda["descricao"]):
                n_legenda_incompleta += 1
            fotos.append({
                "vm": vm, "row": row, "col": col, "_pil": pil_img,
                **legenda,
            })

        plano = composer.plan(taxonomy_key, fotos)
        nome_canonico = _nome_canonico(taxonomy_key)
        img_painel = _desenhar_painel(
            nome_canonico, periodo_txt, fotos,
            n_colunas=plano.n_colunas, coluna_mm=PANEL_WIDTH_MM / plano.n_colunas,
            panel_width_mm=PANEL_WIDTH_MM)

        nome_arquivo = "%s.png" % _slug_arquivo(taxonomy_key)
        caminho_saida = salvar_png(img_painel, out_dir, nome_arquivo)
        altura_mm = img_painel.height / 300.0 * 25.4

        n_sem_imagem = sum(1 for f in fotos if f["_pil"] is None)
        status = "OK"
        notas = []
        if n_legenda_incompleta:
            status = "HITL"
            notas.append("%d de %d fotos sem Estaca/Data/Descricao completos na planilha "
                        "(legenda emitida em branco para essas -- nunca inventada)"
                        % (n_legenda_incompleta, len(fotos)))
        if n_sem_imagem:
            status = "HITL"
            notas.append("%d de %d fotos sem imagem resolvivel na cadeia rich-value"
                        % (n_sem_imagem, len(fotos)))

        provenance = {
            "source_file": caminho, "sheet": aba_nome,
            "method": ("extracao de imagens 'na celula' (Excel rich value: metadata.xml -> "
                      "richData/rdrichvalue.xml -> richValueRel.xml -> media), na ordem das "
                      "ancoras (varredura row-major da planilha); legenda lida por TEXTO DE "
                      "CELULA (rotulos 'Estaca:'/'Data:'/'Coordenada:'/'Descricao:' localizados "
                      "por busca de rotulo, nao por coluna fixa), ja transcrita pela engenharia "
                      "na propria planilha -- sem OCR"),
            "n_fotos": len(fotos),
            "grid": "%d colunas x %d linhas" % (plano.n_colunas, plano.n_linhas),
            "photo_selection_status": "OK" if status == "OK" else "HITL",
        }
        resultados[taxonomy_key] = FiguraResult(
            caminho_saida, PANEL_WIDTH_MM, round(altura_mm, 2), status,
            "; ".join(notas) if notas else None, provenance)

    # servicos do report_data sem aba correspondente encontrada
    encontrados = set(resultados)
    for chave in servicos_relatorio - encontrados:
        resultados[chave] = FiguraResult(
            None, PANEL_WIDTH_MM, 0.0, "HITL",
            "nenhuma aba visivel-com-foto do RELATORIO FOTOGRAFICO mapeia para este servico "
            "(config/services_taxonomy.yaml: abas_relatorio_fotografico)",
            {"source_file": caminho, "photo_selection_status": "NOT_FOUND"})

    return resultados
