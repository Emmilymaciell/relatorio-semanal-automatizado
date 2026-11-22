---
name: session-start
description: Protocolo oficial de ABERTURA de sessão deste projeto. Reconstrói o contexto lendo só o repositório — regras, STATE.md, Git, último checkpoint, estado do shadow report — reconcilia com o disco, mostra divergências e para aguardando autorização. Não altera nada.
when_to_use: No começo de qualquer sessão de trabalho neste repositório, antes de ler documentação ou tocar em arquivo. Também quando o contexto se perdeu (compactação, nova janela, outro computador, outra conta) e é preciso redescobrir onde o projeto parou.
disable-model-invocation: true
disallowed-tools: Edit Write NotebookEdit
argument-hint: "[objetivo da sessão, opcional]"
---

# /session-start — reconstruir o contexto (SOMENTE LEITURA)

A conversa é descartável; **o repositório é a memória**. Esta sessão não herda nada da anterior.

**Proibido aqui:** implementar, editar, criar, refatorar, corrigir "rapidinho", instalar, commitar,
`push`, `stash`, `reset`, apagar, rodar script desconhecido, iniciar a próxima tarefa.
`Edit`/`Write` estão bloqueados pelo frontmatter — se precisar deles, a sessão de trabalho já começou
e não é mais esta skill.

Só inspeção. Comandos de leitura, `git` read-only, e as verificações do passo 5 quando fizerem sentido.

## 1. Onde estou

`pwd`, raiz do Git, branch, HEAD completo e curto, `git status --porcelain`, `git describe --tags`.
Confirme que é **relatorio-semanal-automatizado**. Sem remoto — não procure um.

Separe o status em: staged · não-staged · untracked relevantes. Alterações na árvore **não são sujeira**:
são trabalho de alguém. Nunca sugira descartá-las.

## 2. Regras permanentes (leia primeiro, nesta ordem)

1. `CLAUDE.md` — SharePoint read-only, MAX_PATH, heredoc, níveis de confiança, política de commit
2. `STATE.md` — fase atual, próximo passo exato, pendências, baselines

Pare aqui. Não abra `docs/` inteiro.

## 3. Continuidade

- Último checkpoint em `docs/checkpoints/` (o mais recente por data no conteúdo, não por mtime)
- `docs/checkpoints/SESSOES.md` — histórico compacto das sessões
- `git log --oneline -8`

Abra **um** checkpoint: o que o `STATE.md` indicar para retomada. Os outros 14 são história.

## 4. Reconciliar — evidência vence prosa

Compare `STATE.md` e o checkpoint **contra** `git log`, `git status`, artefatos em disco e
`output/shadow/<obra>_<periodo>/`.

Classifique: concluído · parcial · em andamento · não iniciado.

Se divergirem, **mostre a divergência**; não a corrija em silêncio e não escolha um lado sozinho.
Exemplo real: `STATE.md` pode citar um HEAD que já não é o atual, ou um candidate que a árvore
mostra incompleto.

Estado de processamento a recuperar quando houver (ver [references/estado-processamento.md](references/estado-processamento.md)):
versões do candidate já montadas, semana/obra corrente, corpus já extraído, guard rodado.

## 5. Baseline — separe registrado de executado

`STATE.md` guarda **o último resultado registrado**. Isso não é uma execução de hoje.

Se for útil estabelecer baseline, rode e **diga que rodou agora**:

```bash
for f in tests/test_*.py; do python -X utf8 "$f" >/dev/null 2>&1 || echo "FALHOU: $f"; done
```

Nunca apresente número histórico como recém-medido.

## 6. Próxima unidade de trabalho

A **menor** unidade coerente — não um replanejamento. Objetivo, por quê, critério de pronto,
arquivos prováveis, dependências, riscos, verificações esperadas.

Derive do `STATE.md` → "Próximo passo exato". Se ele estiver vago ou contradito pela árvore, diga isso.

## 7. Antes de perguntar

Procure em: arquivos, Git, `docs/`, `config/*.yaml`, `STATE.md`, checkpoints, `output/`, testes.
Pergunte só quando a resposta **só puder vir da Emmily** — decisão material, risco, ou evidência
insuficiente. Nunca invente para evitar perguntar.

## 8. Relatório

Monte o briefing conforme [references/briefing.md](references/briefing.md) e termine exatamente com:

> Contexto reconstruído. Nenhuma alteração foi realizada.

**PARE.** Não comece a próxima tarefa. Aguarde autorização explícita.
