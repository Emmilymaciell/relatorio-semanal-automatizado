# Critérios de release — "MVP READY FOR SHADOW USE"

Objetivo: definir, de forma **objetiva e verificável**, quando o gerador pode ser usado em
modo sombra (rodando em paralelo ao processo manual, com revisão humana antes de qualquer
uso). "Shadow use" **não** é publicação: nada sai sem a Emmily revisar.

O comando único é `python scripts/generate.py` (ver [`COMO_GERAR.md`](COMO_GERAR.md)). Cada
rodada produz, em `output/shadow/<OBRA>_<ini>_<fim>/`, os artefatos usados abaixo.

---

## 1. Checklist de release

Todos os 9 critérios precisam estar atendidos. Um critério não verificável conta como **não atendido**.

| # | Critério | Como verificar | Arquivo-evidência |
|---|---|---|---|
| R1 | **As 4 obras testadas** (RD-101, RD-102A, RD-103, RD-104), cada uma com ao menos 1 rodada completa pelo comando único | `python scripts/generate.py --road <OBRA> --start <sáb> --end <sex> --versao <v>` para cada obra; exit code 0 | `scorecard_<v>.json` com `veredito = READY_FOR_SHADOW_REVIEW` em cada pasta de obra |
| R2 | **1 semana inédita** — semana v3 **histórica, com golden**, que nunca foi usada para desenvolver o motor (teste de generalização temporal). Reservada: **RD-101 RO_046, 15/08–21/08/2026** (`docs/checkpoints/CURRENT.md`). Gerada do zero, sem adaptação prévia, com `--modo HISTORICAL_SHADOW` | `python scripts/generate.py --road RD101 --start 2026-08-15 --end 2026-08-21 --versao <v>` | `scorecard_<v>.json` com veredito READY; divergências triadas pela taxonomia do §3. *Semana corrente sem golden (`--modo CURRENT`) é uso em sombra propriamente dito, depois do MVP: os gates que comparam com golden precisam virar `N/A` não bloqueante — item pós-MVP.* |
| R3 | **Validador sem FAIL** (15 regras sobre o PDF final montado) | ler o scorecard; ou `python -X utf8 -c "import json;print([r['id'] for r in json.load(open('<arq>',encoding='utf-8'))['resultados'] if r['status']!='PASS'])"` | `validacao_pdf_final_<v>.json` → `scorecard_<v>.json: validador.fail == 0` |
| R4 | **Critical gates 12/12 PASS** | idem | `critical_gates_<v>.json` → `scorecard_<v>.json: gates.pass == gates.total == 12` |
| R5 | **Zero regressão crítica** versus a versão base anterior da mesma obra/semana | `generate.py ... --base-versao <vAnterior>` (ou `python scripts/21_regression_guard.py --obra ... --base ... --nova ...`) | `regression_guard_<base>_<nova>.json` → `veredito` ∈ {`LIBERA`, `LIBERA_COM_JUSTIFICATIVA`}; `regressoes == []`. Exceções justificadas ficam em `config/regression_exceptions.yaml` |
| R6 | **Pipeline reproduzível**: repetir o comando com o mesmo código/config/fontes produz o mesmo resultado (mesmos fatos, contagens, veredito) | rodar duas vezes com versões diferentes e comparar `scorecard_*.json` (campos `validador`, `gates`, `blocos_por_status`, `n_paginas`); conferir `git.head` e `config_sha1` no summary; `pip install -r requirements.txt` em ambiente limpo | `generation_summary_<v>.json` (`git`, `config_sha1`) + `candidate_manifest_<v>.json` (`source_hashes`, `config_hashes`) |
| R7 | **Provenance** completa: todo dado do relatório rastreia até a fonte | arquivo existe e não está vazio; `report_data.json` sem campos órfãos | `report_provenance.json`; `docs/mapeamento/08_data_lineage.md` |
| R8 | **Manifest** íntegro: versão, commit, hashes sha1 das fontes e de `config/*.yaml`, contagem de blocos por status | arquivo existe; `commit` bate com o HEAD usado; `source_hashes` não vazio | `candidate_manifest_<v>.json` |
| R9 | **Comando único documentado e funcional** a partir de um clone limpo, por quem nunca viu o projeto | seguir só o `COMO_GERAR.md` | `docs/COMO_GERAR.md`; exit codes 0/1/2/3 conforme o guia |

Condições transversais (valem para todas as rodadas): nenhum arquivo escrito dentro do SharePoint
(`prog_util.assert_read_only`); suíte `tests/` passando; nenhum DOCX/PDF/PNG versionado.

### Veredito mecânico do scorecard

`READY_FOR_SHADOW_REVIEW` somente se **todas** as condições valem; qualquer uma falhando dá `BLOCKED` (exit 1):

- nenhuma etapa em erro;
- validador presente e **0 FAIL** nas 15 regras;
- critical gates presentes e **todos** `CRITICAL_GATE_PASS` (gate `N/A` também bloqueia);
- se houve `--base-versao`: guard sem regressão (`veredito != BLOQUEIA`).

O veredito READY **não** zera os HITL: pendências humanas ficam listadas no scorecard e precisam ser
resolvidas na revisão.

---

## 2. Quando declarar o MVP

1. R1–R9 marcados com evidência (caminho do arquivo + valor) numa tabela em `STATE.md` ou num
   `docs/checkpoints/CHECKPOINT_*.md`.
2. Revisão humana da Emmily de ao menos 1 relatório por obra (leitura do PDF montado + scorecard).
3. A Emmily declara, por escrito, "MVP pronto para uso em sombra". O comando não declara isso sozinho.

---

## 3. Taxonomia de triagem multiobra

Quando uma obra/semana nova falha (FAIL do validador, gate, regressão ou HITL inesperado), classifique
**antes** de corrigir. A classe define quem corrige e onde a correção mora. Evidência vence prosa:
registre caminho + página/célula + valor.

| Classe | Definição | Exemplo | Quem corrige |
|---|---|---|---|
| `GLOBAL_ENGINE_BUG` | O motor erra de forma independente da obra: lógica, parser, template, regra do validador ou gate. Aparece (ou apareceria) em mais de uma obra. | Título do gráfico de NC fixo em "RD-103" em vez de vir de `report_data` (commit `acee294`); patch do Termo apagando a seção 9 por geometria fixa. | Desenvolvimento do motor (`scripts/gerador/**`, validador). Exige teste de regressão **e** nova rodada nas 4 obras. |
| `LEGITIMATE_ROAD_CONFIG` | A diferença é legítima e específica da obra; o motor deve ser **configurado**, não "corrigido". | Frase-ponte da seção 3 emitida em RD-101/RD-103 e não em RD-102A/RD-104 (`road_registry.s3_frase_ponte_grafico`). | Quem conhece a obra, via `config/road_registry.yaml` / `config/*.yaml`. Nunca via `if obra == ...` no código. |
| `SOURCE_LIMITATION` | A fonte não traz o dado (ou traz em formato inutilizável); o gerador não tem como inventar. Vira `LIMITATION`/`PLACEHOLDER` explícito. | Ficha de Campo ausente em 30/08 → clima `INDETERMINADO`; Dashboard Ambiental é Power BI (sem exportação). | Fora do código: quem produz a fonte (campo/engenharia). O motor só precisa sinalizar com clareza, sem preencher por palpite. |
| `GOLDEN_DEFECT` | O relatório de referência (golden) publicado está errado; o shadow está certo e a divergência é do golden. | Nº de páginas do Termo que não bate com nenhuma contagem; nome da pasta no SharePoint divergente da capa (ver `docs/validacao/CLASSIFICACAO_FALHAS_HISTORICAS.md`). | Registrar como exceção justificada (`config/regression_exceptions.yaml`/comparação) e avisar o autor do relatório; o motor não "copia o erro". |
| `HITL_REQUIRED` | Exige julgamento humano por natureza: texto narrativo, parecer, interferência em campo, conferência de dado ambíguo. Não é bug nem defeito de fonte. | Seções 1 (Apresentação), 5 (Execução) e 9 (Interferência) saem `HITL`. | A Emmily (ou o responsável da obra) na revisão do shadow. O gerador deixa marcado `[HITL]`; nunca resolve sozinho. |
| `VISUAL_POLISH` | Conteúdo e dados corretos; diferença puramente estética (espaçamento, quebra de página, fidelidade tipográfica, pixel-perfect do unifilar). Não afeta o veredito mecânico. | Unifilar renderizado "não pixel-perfect" frente ao Excel; quebra de página diferente do golden (`docs/candidate/PAGINACAO_V0_3.md`: não forçar). | Backlog de acabamento do motor/renderer, priorizado depois dos bloqueios; revisão visual humana decide se importa. |

### Árvore de decisão rápida

1. O dado/fato está **errado** no shadow? → a fonte tem o valor correto? **Sim**: `GLOBAL_ENGINE_BUG` (se afeta qualquer obra) ou `LEGITIMATE_ROAD_CONFIG` (se só essa obra difere por regra). **Não**: `SOURCE_LIMITATION`.
2. O dado do shadow está **certo** e o golden diverge? → `GOLDEN_DEFECT`.
3. Depende de julgamento/redação humana? → `HITL_REQUIRED`.
4. Tudo certo, só a aparência difere? → `VISUAL_POLISH`.
5. Na dúvida entre 1 e 2: rode as outras 3 obras. Se o mesmo sintoma aparece (ou aparecerá) nelas, é `GLOBAL_ENGINE_BUG`.

Todo `GLOBAL_ENGINE_BUG` corrigido exige: teste automatizado em `tests/`, rodada nas 4 obras e guard sem regressão.
