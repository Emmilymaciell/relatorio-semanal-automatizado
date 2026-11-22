# Relatório de encerramento — modelo

Carregue no passo 7 do `SKILL.md`. Omita campo inaplicável; não preencha por preencher.

---

```
SESSÃO
Número:
Nome:
Objetivo:

REPOSITÓRIO
Diretório:
Branch:
HEAD:            <curto> — <mensagem>
Working tree:    limpa | N staged, N não-staged, N untracked

CONCLUÍDO
PARCIALMENTE CONCLUÍDO
INVESTIGADO / NÃO IMPLEMENTADO
    (o que foi estudado e deliberadamente não virou código — e por quê)

DESCOBERTAS IMPORTANTES
DECISÕES DURÁVEIS            (DEC-NNN criadas, se houver)
HIPÓTESES / QUESTÕES ABERTAS

VALIDAÇÕES EXECUTADAS
Comando:
Resultado:
Não executado:               (o quê, por quê, risco residual)

ALTERAÇÕES NÃO COMMITADAS
    (listar; separar o que é desta sessão do que já estava na árvore antes)

DOCUMENTAÇÃO ATUALIZADA
    (arquivo → o que mudou)

ESTADO DE PROCESSAMENTO
    (obra/semana, versões do candidate, último guard — quando aplicável)

INCONSISTÊNCIAS ENCONTRADAS

PRÓXIMO PASSO EXATO
Objetivo:
Critério de pronto:
Dependências:
Riscos:
Verificações esperadas:

ARQUIVOS DE ALTO SINAL PARA A PRÓXIMA SESSÃO
    (no máximo 5)

RISCOS / BLOQUEIOS

GIT FINAL
Branch:
HEAD:
Status:
Commit:          não feito — política pede pedido explícito | <hash>, a pedido da Emmily
```

> Handoff concluído. Nenhuma nova implementação foi iniciada.

---

## Cuidados ao preencher

**Separe o que é seu do que já estava.** A árvore pode conter trabalho anterior não commitado.
Misturar os dois na mesma lista faz a próxima sessão commitar o que não devia.

**Validação.** Comando exato e resultado real. `19/24` é informação; "quase tudo passou" não é.
Se não rodou, diga que não rodou.

**Commit.** A política é pedido explícito (`CLAUDE.md`). Terminar a sessão não é pedido.
Ofereça em uma linha; não commite para "deixar organizado".

**Erro cometido na sessão.** Registre o raciocínio errado, não só a correção — é o que impede
alguém (inclusive você) de repeti-lo. Exemplo que já aconteceu aqui: concluir que um arquivo
estava desatualizado pelo `mtime`, quando o `mtime` era só a data da cópia.

**Próximo passo.** Executável. "Continuar a comparação" não serve; "rodar
`scripts/22_comparacao_exaustiva_v0_2.py` e reavaliar DADO-09" serve.
