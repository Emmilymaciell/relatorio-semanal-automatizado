# Logos institucionais do cabeçalho

> **Versão pública:** os dois PNG desta pasta são placeholders genéricos (retângulo com o texto
> "PARCEIRO" / "PROGRAMA") nas mesmas dimensões dos originais; os logotipos reais foram removidos.
> A seção "Procedência" descreve de onde os originais foram extraídos no projeto de origem.

| Arquivo | Pixels | Proporção | Uso no cabeçalho | Tamanho impresso |
|---|---|---|---|---|
| `logo_parceiro.png` | 243 × 81 | 3,00 | **esquerda**, encostado na margem | **41,1 × 13,7 mm** |
| `logo_prog.png` | 129 × 69 | 1,87 | **direita**, encostado no fim da área útil | **21,9 × 11,7 mm** |

As proporções impressas batem exatamente com as dos arquivos — os logos são
colocados no aspecto natural, sem distorção.

## Procedência

Extraídos de `word/media/` de
`RD103\…\2026.08.22-2026.08.28 48º Relatório Semanal\RO_048.1 - RD-103.docx`
(**semana 48**), em 2026-09-17.

**Deliberadamente não vieram do golden em comparação** (`RO_048` ≠ `RO_049.1`), pela
mesma regra dos textos fixos: ativo de template do processo pode ser derivado do corpus,
desde que **não** do relatório que está sendo usado como referência de comparação.
Os três cabeçalhos do documento (`header1`, `header2`, `header3`) referenciam **as mesmas
duas imagens** — são ativos institucionais, não conteúdo da semana.

O SharePoint não foi alterado: o `.docx` foi copiado para `tmp/` antes de ser aberto.

## Posicionamento

**Não usar coordenada absoluta.** O logo da direita se ancora no fim da área útil
da seção:

```
usable_width = page_width − left_margin − right_margin
x_logo_prog  = usable_width − 21,9 mm
comprimento_da_linha = usable_width
```

Em A4 do corpo (160 mm úteis) isso dá 138,1 mm — bate com os 137,7 mm medidos no
`header3` do golden. No A3 do fotográfico (247 mm úteis) a mesma fórmula resolve sozinha.

É exatamente por **não** seguir essa regra que o golden tem o defeito do `header2`:
coordenada fixa de A3 (228,7 mm) herdada em páginas A4, jogando o logo para fora da
folha nas páginas 4, 5 e 6.
