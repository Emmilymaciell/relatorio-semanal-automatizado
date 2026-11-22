# Como gerar um Relatório Semanal (comando único)

Guia para quem nunca viu o projeto. O comando gera o relatório-sombra de **uma obra, uma semana**,
converte para PDF, monta o PDF final (anexos + Termo), roda o validador e os critical gates e emite
um scorecard. **Nada é publicado automaticamente** — o resultado vai sempre para revisão humana.

## Pré-requisitos

- Windows 11 com **Microsoft Word instalado** (a conversão DOCX→PDF usa o Word via COM).
- **Python 3.12** e os pacotes do projeto: `python -m pip install -r requirements.txt`
  (python-docx, PyMuPDF, openpyxl, PyYAML, pillow, reportlab, pywin32, lxml).
- **Acesso de leitura ao SharePoint sincronizado** em
  `C:\Users\<usuário>\PastaSincronizada\` (OneDrive com os arquivos baixados localmente).
  O projeto **nunca escreve** nessa pasta.
- Este repositório clonado. O caminho do repositório é descoberto sozinho (`prog_util.PROJ`, derivado
  de `__file__`); o SharePoint e o `tmp/` podem ser redirecionados pelas variáveis de ambiente
  `PROG_SHAREPOINT` e `PROG_TMP` (requer o `prog_util.py` da branch principal, commit `7dcca07`).
- Primeira execução numa máquina/clone novo é **mais lenta**: a pasta `tmp/` (cópias de arquivos do
  SharePoint) começa vazia e é repovoada automaticamente (planilhas de ~160 MB).

## Comando

```powershell
python scripts/generate.py --road RD103 --start 2026-08-29 --end 2026-09-04 --versao mr1
```

| Opção | Para quê |
|---|---|
| `--road` | `RD101`, `RD102A`, `RD103` ou `RD104` (aceita `RD-103`) |
| `--start` / `--end` | sábado e sexta da semana (`--end` é opcional: início + 6 dias) |
| `--versao` | rótulo da rodada (nomeia todos os arquivos de saída). Default `mr1` |
| `--modo` | `HISTORICAL_SHADOW` (default; semana já emitida) ou `CURRENT` (semana em curso/inédita) |
| `--base-versao v0.2` | roda também o **regression guard** contra essa versão anterior |
| `--pdf-manual` | para após o DOCX e imprime como converter no Word à mão |
| `--skip-pdf` | usa um PDF já existente em vez de converter |
| `--resume-from assemble` | retoma na montagem (depois de converter o PDF à mão) |
| `--fotos`, `--extrair-paginas` | conferência extra de fotos no guard / PNG de cada página |

Duração típica: alguns minutos (a conversão do Word leva ~45 s; duas conversões simultâneas
são enfileiradas por um lock).

## O que sai

Tudo em `output/shadow/<OBRA>_<início>_<fim>/` (dentro do repositório):

| Arquivo | O que é |
|---|---|
| `shadow_final_candidate_<v>.docx` | relatório editável |
| `shadow_final_candidate_<v>_MONTADO.pdf` | **PDF final**, com anexos e Termo |
| `SCORECARD_<v>.md` / `scorecard_<v>.json` | veredito e checagens (leia este primeiro) |
| `generation_summary_<v>.json` | comando, HEAD do git, sha1 da config, status de cada etapa, caminhos |
| `validacao_pdf_final_<v>.json` | 15 regras do validador |
| `critical_gates_<v>.json` | 12 critical gates |
| `candidate_manifest_<v>.json`, `report_provenance.json`, `report_data.json` | manifesto, proveniência e dados |
| `regression_guard_<base>_<v>.json` | só com `--base-versao` |
| `generate_<v>.log` | log completo |

DOCX e PDF **não vão para o Git** (contêm fotos da obra; estão no `.gitignore`).

## Códigos de saída

| Código | Significado |
|---|---|
| 0 | `READY_FOR_SHADOW_REVIEW` — pronto para **revisão humana** |
| 1 | `BLOCKED` — algum FAIL do validador, gate não-PASS, etapa com erro ou regressão. Veja os motivos no fim da execução e no scorecard |
| 2 | argumentos inválidos |
| 3 | PDF manual pendente (conversão automática desativada ou falhou) |

## PDF manual (quando o Word falha)

O comando imprime o DOCX a abrir e o caminho do PDF a salvar. Abra o DOCX no Word, **aceite a
atualização de campos**, salve como PDF exatamente no caminho indicado e rode de novo com
`--resume-from assemble` (mesma `--versao`).

## O que fazer com HITL / LIMITATION / PLACEHOLDER

O scorecard lista as pendências. `READY_FOR_SHADOW_REVIEW` **não** as elimina:

- `HITL` — texto/decisão que exige humano (ex.: Apresentação, Execução, Interferência em campo).
  Revise e preencha no DOCX.
- `LIMITATION` — a fonte não traz o dado (ex.: Dashboard Ambiental em Power BI). Complete à mão ou
  registre como limitação.
- `PLACEHOLDER` — espaço reservado (ex.: Termo, que é fechado na montagem).

Classifique falhas inesperadas pela taxonomia de [`RELEASE_CRITERIA.md`](RELEASE_CRITERIA.md).

## Garantias

- O SharePoint é somente leitura; toda escrita fica neste repositório.
- Nenhum envio, upload, e-mail ou publicação. Commits só quando a Emmily pedir.
- O gerador é determinístico dado o mesmo código, config e fontes (`config_sha1` e `git.head` ficam no
  `generation_summary`).
