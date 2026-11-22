# Programa — Automação do Relatório Semanal de Acompanhamento de Obras

Engenharia reversa, mapeamento de dados e (futuramente) automação dos
**Relatórios Semanais de Acompanhamento de Obras** do contrato Programa —
obras RD-101, RD-102A, RD-103 e RD-104.

> **Fase atual: DESCOBERTA.** Este repositório ainda não gera relatórios.
> Ele documenta *como os relatórios são feitos hoje* e *de onde vem cada dado*.

---

## ⛔ O SharePoint é read-only

`C:\caminho\exemplo\PastaSincronizada\` é fonte de informação e **nunca**
é modificado por este projeto. Toda escrita acontece dentro deste repositório.
Ver `CLAUDE.md`.

---

## Estrutura

```
relatorio-semanal-automatizado/
├── config/        # conhecimento machine-readable (YAML) — schema, fontes, regras
├── docs/
│   ├── arquitetura/  # arquitetura proposta da automação
│   ├── mapeamento/   # documentação viva da engenharia reversa (01..14)
│   ├── relatorios/   # análises por relatório individual
│   ├── fontes/       # fichas dos documentos-fonte
│   └── decisoes/     # DEC-001, DEC-002, ... (decisões arquiteturais)
├── examples/      # trechos ilustrativos extraídos (texto, nunca originais)
├── logs/          # logs de execução dos scripts
├── output/
│   ├── inventarios/  # inventários do ecossistema documental
│   ├── comparacoes/  # diffs semânticos entre relatórios
│   ├── evidencias/   # evidence ledger (valor ↔ arquivo ↔ local)
│   └── mapas/        # grafos e mapas de dependência
├── scripts/       # ferramentas de investigação (numeradas)
├── state/         # estado intermediário das varreduras
└── tests/         # testes das regras descobertas
```

---

## Documentos principais

| Arquivo | Conteúdo |
|---|---|
| `STATE.md` | Onde a investigação está agora. **Leia primeiro.** |
| `CLAUDE.md` | Regras operacionais (read-only, MAX_PATH, confiança). |
| `docs/RELATORIO_ENGENHARIA_REVERSA.md` | Documento central de descoberta (produzido ao fim da fase). |
| `docs/mapeamento/` | Documentação viva (01 a 14). |
| `docs/mapeamento/14_perguntas_para_emmily.md` | As 11 perguntas que não deram para resolver nos documentos. |
| `config/report_schema.yaml` | Modelo canônico do relatório + 7 regras de validação. |
| `config/road_registry.yaml` | Cadastro mestre das 4 obras. |

---

## Scripts

Requerem Python 3.12 com `pymupdf`, `python-docx`, `openpyxl`, `pdfplumber`.

```bash
python scripts/01_inventariar_relatorios.py
```

| Script | Função |
|---|---|
| `prog_util.py` | Utilitários comuns: long paths, cópia segura, normalização. |
| `01_inventariar_relatorios.py` | Inventaria as 193 pastas-semana das 4 obras. |
| `02_selecionar_amostra.py` | Seleciona a amostra e detecta regimes de modelo. |
| `03_extrair_pdf.py` | Estrutura dos PDFs: páginas, fontes, imagens, formatos. |
| `04_extrair_docx.py` | XML interno dos DOCX: estilos, seções, campos, mídia. |
| `05_inventariar_excel.py` | OOXML das planilhas (trata arquivos de até 167 MB). |
| `06_inventariar_ecossistema.py` | Inventaria as pastas irmãs (3.944 arquivos). |
| `07_estrutura_e_diff.py` | Seções, legendas, anexos e diff semântico. |
| `08_historico_anexos.py` | Periodicidade dos anexos ao longo do histórico. |

### Testes

```bash
python tests/test_regras_periodo.py   # 144/144 — período e numeração
python tests/test_regras_anexos.py    # 52/52  — composição dos anexos
```

---

## Obras

| Pasta | Rótulo | Semanas inventariadas |
|---|---|---|
| `RD101` | RD-101 | 49 |
| `RD102A` | RD-102A | 50 |
| `RD103` | RD-103 | 50 |
| `RD104` | RD-104 | 44 |
