# Briefing de abertura — modelo

Carregue este arquivo no passo 8 do `SKILL.md`. Omita campo que não se aplique à sessão;
não invente conteúdo para preencher seção.

Escreva em português do Brasil, direto. Números verificados, não arredondados de memória.

---

```
PROJETO
Diretório:
Objetivo geral:

SESSÃO
Número:            (próximo em docs/checkpoints/SESSOES.md)
Nome sugerido:     N - <nome objetivo>
Objetivo desta sessão:

GIT
Branch:
HEAD:              <curto> — <mensagem>
Tag mais recente:
Status:            limpo | N staged, N não-staged, N untracked

ÚLTIMO ESTADO VERIFICÁVEL
(o que o repositório prova, não o que o STATE.md narra)

CONCLUÍDO
PARCIALMENTE CONCLUÍDO
EM ANDAMENTO
PENDENTE

ESTADO DE PROCESSAMENTO
(quando houver: obra/semana corrente, versões do candidate, corpus, guard)

PRÓXIMA TAREFA EXATA
Objetivo:
Critério de pronto:
Arquivos prováveis:
Dependências:
Riscos:
Verificações esperadas:

ARQUIVOS DE ALTO SINAL
(no máximo 5 — o que realmente precisa ser lido antes de agir)

DECISÕES RELEVANTES
(DEC-NNN que restringem o que vem a seguir)

VALIDAÇÃO
Último resultado registrado:   <de onde saiu>
Verificações executadas agora: <comando + resultado real, ou "nenhuma">

PENDÊNCIAS / BLOQUEIOS

INCONSISTÊNCIAS ENCONTRADAS
(STATE.md × Git × disco — mostrar, não resolver)
```

> Contexto reconstruído. Nenhuma alteração foi realizada.

---

## Cuidados ao preencher

**Números.** Conte, não estime. Se disser "24 suítes", foi `ls tests/test_*.py | wc -l`.

**Validação.** Nunca apresente resultado registrado no `STATE.md` como execução de hoje.
Se não rodou nada, escreva `nenhuma`.

**Níveis de confiança.** O projeto usa `CONFIRMADO` · `ALTA CONFIANÇA` · `PROVÁVEL` · `HIPÓTESE` ·
`DESCONHECIDO` (`CLAUDE.md`). Preserve o nível de origem: hipótese não vira fato ao ser repetida
num briefing.

**Árvore suja.** Alterações não commitadas são trabalho, não lixo. Liste-as e siga. Não proponha
`stash`, `reset` nem `checkout --`.

**Sessão sem número claro.** Se `SESSOES.md` não permitir deduzir o número, diga que não é possível
determinar e sugira um nome sem número. Não invente numeração.
