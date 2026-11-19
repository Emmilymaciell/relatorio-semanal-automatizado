# Fase 3.3 — Final Candidate v0.1

**Baseline:** `ff80d95` · **Criado em:** 2026-09-17
**Meta 3.3A:** `output/shadow/RD103_2026-08-29_2026-09-04/shadow_final_candidate_v0.1.docx`
**sem placeholder relevante.**

Autoridade visual: `docs/visual_review/GOLDEN_VISUAL_SPEC.md`.
**Não fazer nova engenharia reversa.** **Golden nunca é fonte de dado.**

---

## 0. Regras que não mudam

- SharePoint **READ-ONLY**; `prog_util.lp()` / `assert_read_only()`.
- **Não sobrescrever** `shadow_report.docx` nem `shadow_report_v2.docx`.
- **Zero IA/LLM** no código. **Nada de `if obra == "RD103"`.**
- `.py` com Write/Edit, nunca heredoc do Bash.
- Word/Excel COM **não funcionam** aqui. Não insistir. Tudo que depender deles é
  substituído por renderização nativa.
- Sem commit, push ou tag pelos agentes.

## 0.1 Bibliotecas disponíveis — não instalar nada

`PIL` 12.1 · `numpy` · `fitz` (PyMuPDF) · `python-docx` · `openpyxl` · `pypdf` ·
`reportlab` · `pdfplumber`.
**`matplotlib` e `qrcode` NÃO existem** — gráficos se desenham com **PIL**.

---

## 1. Contrato comum das figuras

Todo módulo de figura vive em `scripts/gerador/figuras/` e devolve:

```python
@dataclass
class FiguraResult:
    path: str | None          # PNG gerado em output/shadow/<...>/figuras/
    largura_mm: float         # tamanho de inserção no Word
    altura_mm: float
    status: str               # OK | LIMITATION | PRESENTATION_DEPENDENCY | HITL
    note: str | None
    provenance: dict          # {source_file, sheet, cells|range, method}
```

**Resolução:** gerar o PNG a **300 dpi** sobre a dimensão em mm
(`px = mm / 25.4 * 300`), para o PDF sair limpo.

O renderer insere `FiguraResult.path` com `largura_mm × altura_mm`. Se `path is None`,
mantém o placeholder **com o motivo**.

---

## 2. Figuras a produzir

### 2.1 `Figura 1` — Ficha de Acompanhamento (`figuras/ficha_acumulada.py`)

Fonte: `Arquivos Auxiliares/Ficha de Campo Acumulada RD-103.xlsx`, aba `FICHA CAMPO`.
Desenhar **nativamente com PIL** a tabela de estacas (cabeçalho + linhas), no estilo da
seção 5 do `GOLDEN_VISUAL_SPEC.md`: bordas finas, cabeçalho sombreado, Arial.
Largura = largura útil A4 (**160 mm**).

### 2.2 `Gráfico 1` — Avanço Global (`figuras/avanco_grafico.py`)

Fonte: `progress.series` do `report_data.json` (aba `Relatório Semanal` do
`Gráfico - Avanço Global.xlsx`) — **já extraída**, 4 semanas.
Barras agrupadas por semana: `Concluído` e `Concluído + Atacado`, rótulo de valor,
eixo de datas. **PIL**, não matplotlib. **Não copiar imagem do golden.**
Valores desta semana: **32,23 %** e **42,19 %**.

### 2.3 Relatório Fotográfico (`figuras/painel_fotografico.py`) — **o item central**

> **Fonte de seleção — resolvida.** `Arquivos Auxiliares/RELATORIO FOTOGRAFICO 49 - 180.xlsx`
> contém **63 imagens embutidas** (45 PNG + 17 JPEG) distribuídas nas abas
> `TERRAPLENAGEM · PAVIMENTAÇÃO · OAC · DRENAGEM · OAE · CONSERVAÇÃO · MEIO AMBIENTE · ADM`.
> **São as fotos já escolhidas pela engenharia** — fonte legítima do processo,
> não o relatório final. Isso resolve a seleção sem tocar no golden.

Leitura: abrir o `.xlsx` como zip, casar `xl/drawings/drawingN.xml` → aba pelo
`xl/worksheets/_rels/`, e extrair as imagens **na ordem das âncoras** (ordem visual).

Composição por serviço, conforme `GOLDEN_VISUAL_SPEC.md` §7:

| | |
|---|---|
| Papel | A3 retrato, largura útil **247 mm** |
| Grid | **3 colunas fixas**, linhas = `ceil(n/3)` |
| Coluna | ≈ **80 mm** |
| Altura por foto | 50–57 mm conforme o aspecto |
| Legenda | **3 linhas acima** de cada foto: `Estaca + Data` / `Coordenada` / `Descrição` |
| Linha incompleta | células vazias **com borda**; não expandir as fotos |

Os dados da legenda vêm do **overlay já impresso na foto** (não há OCR: o overlay está
dentro do pixel) cruzado com a **Ficha de Campo** do dia. Se não for possível cruzar com
segurança, **emitir a foto sem a legenda textual** e marcar o serviço como
`photo_selection_status: HITL` — **nunca inventar estaca, data ou coordenada**.

`photo_selection_status` por serviço: `OK` (aba correspondente encontrada) ·
`HITL` (ambiguidade) · `NOT_FOUND`.

### 2.4 Painéis ambientais (`figuras/ambiental.py`)

Fonte: `04_Acompanhamento - Licenca Ambiental.xlsx` (abas `Licenças`, `Condicionantes`).
Desenhar nativamente os dois painéis com PIL, a partir dos dados estruturados.
Largura **160 mm**; alturas do `GOLDEN_VISUAL_SPEC.md` §6 (86,89 e 63,98 mm).

**A terceira figura é um QR Code.** A biblioteca `qrcode` não existe e **não se instala
nada**. Emitir com `status: PRESENTATION_DEPENDENCY`, quadrado **79,07 × 79,07 mm**,
com o motivo registrado. **Incluir o elemento no relatório**, marcado.

### 2.5 NC — figura e 2 gráficos (`figuras/nc_graficos.py`)

Fonte: `02_Acompanhamento de NC automatizada.xlsm`, `VISAO_GERAL!Tbl_VISAO`, filtrada
pela rodovia. Figura de notificações (160 × 63,82 mm) + `Gráfico 2` por status
(120,82 × 61,48 mm, centralizado) + `Gráfico 3` por natureza (111,25 × 67,37 mm).
Pizza/barras com **PIL**.

Manter a divergência temporal como **`SOURCE_LIMITATION`** (a planilha usa `TODAY()`).
**Nunca forçar o valor do golden.**

---

## 3. As 3 lacunas do Checkpoint 3.2B

1. **Legendas ausentes.** Emitir `Figura 1`, `Gráfico 1` e `Tabela 1`. A sequência não
   pode começar em `Figura 2`. Sequências independentes por tipo.
2. **Estilo `Caption` órfão.** Hoje é 9 pt azul `4F81BD` herdado do template e ninguém
   usa. Corrigir para o padrão real (**Arial 10 pt `0D0D0D` centralizado**) e passar as
   legendas a **usar o estilo**, em vez de formatação direta.
3. **Tabela de NC.** Bordas 0,5 pt, cabeçalho sombreado `F2F2F2`, Arial, larguras
   coerentes — o mesmo tratamento da tabela do clima.

---

## 4. Anexos embutidos (`EMBEDDED_IMAGE`)

PA-254 (6 pág.), Acompanhamento Ambiental (2), SST (2) = **10 páginas**.
Renderizar cada página do PDF-fonte com `fitz` a 200 dpi e inserir **em página cheia**
no DOCX, herdando cabeçalho e numeração, na ordem do `assembly_plan`.
Uma página do documento por página do anexo.

Anexos `MERGED_PDF` (unifilar, diário) **continuam fora do Word** — não transformar em
seção artificial.

---

## 5. Termo de Encerramento — duas passagens

```
PASS 1  monta DOCX -> PDF do corpo -> + PDFs externos -> conta páginas reais
PASS 2  atualiza o Termo com o número -> re-renderiza -> confere
```

Se a `PASS 2` mudar a paginação, permitir **uma** reverificação e parar.
**Máximo 3 passagens**; se não estabilizar, registrar `PAGINATION_UNSTABLE` e seguir.
Enquanto não houver PDF, o Termo fica `PLACEHOLDER` com `n_paginas_provisorio: true`.

---

## 6. Manifesto de versão

`output/shadow/<...>/candidate_manifest_v0.1.json`:

```json
{"candidate_version":"v0.1","commit":"...","generated_at":"...","road":"RD-103",
 "period":{"start":"2026-08-29","end":"2026-09-04"},
 "source_hashes":{"<caminho>":"sha1"},
 "config_hashes":{"visual_style_v3.yaml":"sha1"},
 "renderer_version":"v2","validation_result":{...}}
```

Um PDF tem de poder ser reproduzido a partir dele.

---

## 7. Pipeline

```
canonical -> content_builder -> document_model -> docx_renderer_v2
   -> DOCX -> (PDF do corpo) -> annex assembly -> merge -> PDF completo
   -> Termo/paginação -> validador final
```

CLI: `python scripts/shadow_report.py candidate --obra RD103 --inicio ... --versao v0.1`.
Preparado para as 4 obras com o mesmo código — **não executar a bateria multiobra agora**.

---

## 8. Parada da 3.3A

Quando o DOCX estiver completo **sem placeholder relevante**, parar e pedir a conversão
manual:

```
ABRA:        output\shadow\RD103_2026-08-29_2026-09-04\shadow_final_candidate_v0.1.docx
SALVE COMO:  output\shadow\RD103_2026-08-29_2026-09-04\shadow_final_candidate_v0.1.pdf
```

Não transformar a conversão em bloqueio arquitetural: tudo que independe dela já tem de
estar pronto.

## 9. Pragmatismo

Corrigir P2/P3 que sejam **repetitivos, de impacto visual claro e corrigíveis
globalmente**. **Não parar** por 1 mm, 1 pt ou espaçamento isolado — isso fica para
depois da primeira comparação completa.

## 10. Regressão

As **17 suítes** continuam passando. Visual não pode quebrar conteúdo: período, nº 49,
serviços, avanço 32,23/42,19 %, clima com 30/08 = `-` + HITL, PA-254, NC, anexos.
