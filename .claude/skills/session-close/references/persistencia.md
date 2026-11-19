# Onde e como persistir — modelos

Carregue no passo 3 do `SKILL.md`, só para o destino que a sessão realmente exigir.

---

## `STATE.md` — snapshot, nunca diário

**Substitua** as seções; não acumule. Se uma informação deixou de ser verdade, ela sai — não vira
"histórico" dentro do STATE.

Seções que precisam continuar verdadeiras:

- cabeçalho: data, fase atual, HEAD, tag, `git status`
- `## RETOMADA` — os passos para voltar sem conversa nenhuma
- `## Fases` — tabela de situação
- `## O que está em andamento` — vazio é resposta legítima ("Nada.")
- `## Próximo passo exato` — **uma** ação executável, com comando quando houver
- `## Pendências reais` e `## Perguntas abertas`
- `## Comandos principais` — se mudou, corrija

Remova o que ficou obsoleto. Um `STATE.md` que mente é pior que um desatualizado declarado.

---

## `docs/checkpoints/CHECKPOINT_<NOME>.md`

Só quando houver **marco** (rodada fechada, fase concluída) ou **risco de perda** (contexto no
limite, trabalho grande pela metade).

O contrato, do README de lá: permitir retomar **sem acesso a nenhuma conversa**.

```markdown
# CHECKPOINT <nome> — <uma linha>

**Data:** AAAA-MM-DD · **Fase:** · **HEAD:** <curto>

## RETOMADA
<os passos exatos, em ordem, para continuar>

## Onde paramos
<e o motivo, se parou por algo específico>

## O que esta rodada entregou
<tabela item → situação>

## O que NÃO refazer
<para a próxima sessão não repetir trabalho caro>

## Erros corrigidos no caminho
<o raciocínio errado, para não voltar>

## Artefatos
<caminhos + como foram verificados>

## Próximo passo
```

Acrescente a linha correspondente na tabela de `docs/checkpoints/README.md`.

---

## `docs/checkpoints/SESSOES.md` — uma linha por sessão

Compacto por decisão. Existe porque a política deste projeto proíbe commit automático: sessão
que não commitou não deixaria rastro nenhum.

```markdown
| N | Data | Nome | Entregue | Validação | HEAD final | Próximo passo |
```

Entregue = resultado, não narrativa. Validação = o que rodou de fato.

---

## `docs/decisoes/DEC-NNN-<slug>.md`

Só decisão **durável** — que restringe trabalho futuro. Numeração sequencial a partir da última
em `docs/decisoes/`.

```markdown
# DEC-NNN — <decisão em uma linha>

**Data:** · **Status:** ACEITA | REVOGADA (ver DEC-MMM)

## Contexto
## Evidência
<com nível de confiança: CONFIRMADO / ALTA CONFIANÇA / PROVÁVEL / HIPÓTESE>
## Decisão
## Alternativas consideradas
## Consequências
```

Decisão revogada **não se apaga**: marca-se o status e aponta-se a substituta. É assim que a
DEC-002 convive com a DEC-004 no projeto G2.

---

## Classificação — não promova nada por descuido

| Ficou assim | Registre assim |
|---|---|
| verificado no arquivo, com citação | `CONFIRMADO` |
| várias evidências convergentes | `ALTA CONFIANÇA` |
| uma evidência boa | `PROVÁVEL` |
| plausível, não verificado | `HIPÓTESE` |
| não descoberto | `DESCONHECIDO` |

Registrar como desconhecido é preferível a inventar (`CLAUDE.md`).

Hipótese não vira fato por ter sido útil. Investigação não vira implementação por ter sido
trabalhosa. Teste histórico não vira teste atual por ter passado antes.
