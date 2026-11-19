# DEC-004 — Formato das imagens compostas (Relatório Fotográfico e painéis)

**Data:** 2026-09-17 · **Situação:** ✅ **decidida** · **Fase:** 3.2B (P0.5)
**Resolve:** P0.5 de `docs/arquitetura/FASE3_2B_RENDERER_V2_SPEC.md`

---

## Contexto

O golden insere o Relatório Fotográfico (seção 5.2) e os painéis ambientais como
**EMF** (Enhanced Metafile) — um `wp:inline` por painel, sem tabela nativa
(`VISUAL_DIFF_MASTER.md` §C3, `GOLDEN_VISUAL_SPEC.md` §7). Isso é **consequência do
processo manual**: quem monta o relatório hoje monta o painel no Excel/PowerPoint e
**cola** no Word — colar do Office produz EMF por padrão. Não é um requisito de
qualidade do documento, é um artefato do fluxo humano que a automação está
substituindo.

A dúvida a resolver: o `PhotoPanelComposer` (P0.4) e qualquer futuro compositor de
painel/gráfico **precisam** produzir EMF para o documento final ficar equivalente ao
golden?

---

## Decisão

**EMF não é obrigatório.** A automação pode produzir a imagem composta em **PNG de
alta resolução** (formato padrão desta rodada) **ou SVG**, desde que a imagem
gerada cumpra, cumulativamente:

1. **Fidelidade visual** — mesma composição (grid 3 colunas, molduras, legendas
   acima de cada foto) e mesma leitura que o painel EMF do golden produziria.
2. **Dimensão correta em milímetros** — a imagem inserida no Word precisa
   corresponder exatamente às medidas do painel especificadas em
   `GOLDEN_VISUAL_SPEC.md` §7 (247,00 × 306,32 mm predominante), não um tamanho
   arbitrário de pixel reamostrado na inserção.
3. **Estabilidade no Word** — a imagem não pode "quebrar" ao reabrir o `.docx`
   (recompressão automática do Word, texto de moldura reflowando, etc.) — o que EMF
   evita por ser vetorial, mas que PNG em resolução suficiente e SVG também evitam
   na prática, desde que a imagem seja **inserida com `width`/`height` explícitos em
   EMU** (não deixada para o Word inferir do DPI do arquivo).
4. **Boa qualidade no PDF** — ao exportar DOCX → PDF, a imagem não pode pixelizar
   visivelmente em impressão/zoom razoável.

Formato **default: PNG** (`photo_panel.PhotoPanelResult.image_format = "PNG"`).
SVG fica como alternativa aceitável quando o compositor de imagem futuro preferir
vetor (ex.: reaproveitar as molduras e legendas como texto pesquisável) — mas não é
obrigatório para esta decisão.

**EMF continua PERMITIDO**, caso um caminho de implementação futura opte por gerar
via COM do Office (que produz EMF nativamente) — a decisão não proíbe EMF, apenas
remove a obrigação. O critério de aceite é a lista de 4 itens acima, não o formato do
arquivo.

---

## Justificativa

Reproduzir EMF programaticamente (fora do COM do Office) tem custo alto e baixo
retorno: é um formato de metarquivo do Windows, mal suportado por bibliotecas Python
de geração de imagem, e cuja única vantagem real (ser vetorial/editável no Word) não
é um requisito do relatório — ninguém edita o painel fotográfico depois de gerado.
PNG em resolução alta (300+ DPI no tamanho final em mm) entrega os 4 critérios acima
com ferramentas padrão (Pillow, matplotlib, ou renderização de HTML/SVG para
raster), sem dependência de COM/Office — o que também evita o mesmo tipo de
travamento já observado com a conversão DOCX → PDF via Word neste ambiente.

---

## Consequências

- `scripts/gerador/photo_panel.py` (`PhotoPanelComposer`, `PhotoPanelResult`)
  usa `image_format="PNG"` como padrão; a interface não impõe EMF em nenhum ponto.
- A implementação desta rodada (P0.4) ainda não produz a imagem composta de verdade
  — devolve um **plano de grid** (colunas/linhas/dimensões) e um placeholder pelo
  componente certo. Quando a composição real for implementada (próxima rodada),
  ela deve gerar PNG/SVG nas dimensões em mm da especificação, não herdar nenhuma
  lógica de EMF.
- Qualquer comparação golden × shadow por **conteúdo visual** (não por formato de
  arquivo) deve validar os 4 critérios desta decisão, não "é EMF?".
- Não fecha a porta para EMF via COM se uma rodada futura decidir esse caminho por
  outro motivo (ex.: reaproveitar um painel já montado no Excel) — só deixa de ser
  obrigatório.
