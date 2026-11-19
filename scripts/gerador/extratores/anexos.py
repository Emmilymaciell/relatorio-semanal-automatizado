# -*- coding: utf-8 -*-
r"""
scripts/gerador/extratores/anexos.py - lista de anexos candidatos (regras de
report_schema.yaml s11_anexos), ordem: condicionais primeiro, depois os dois
fixos (Diagrama Unifilar, Diario de Obras), nessa ordem, sempre por ultimo.

Nao produz o PDF do anexo (fora de escopo desta fase) -- so decide QUAIS
candidatos existem e aponta para o arquivo-fonte real, quando localizavel.

Este modulo e' so CONTEUDO (quais anexos existem). COMO cada anexo entra no
arquivo final (EMBEDDED_IMAGE dentro do .docx x MERGED_PDF depois da
exportacao) e' decisao da camada de MONTAGEM (P0.2,
`gerador.assembly_model.classificar_anexo`), lida de
`config/annex_types.yaml` -- nunca por posicao/numero aqui. `montar()`
valida, ao final, que todo `id` que devolve tem classificacao configurada
(falha alto em vez de deixar um anexo sem integration_mode passar batido).
"""
from .. import assembly_model as _assembly


def montar(*, meetings, ambiental_applies, fa_ambiental, fa_sst,
           nc_no_periodo, unifilar_path, unifilar_exists,
           diario_path, diario_exists, parecer=None):
    anexos = []

    if meetings:
        anexos.append({
            "id": "ata_reuniao", "obrigatorio": False, "presente": True,
            "itens": [m.get("path") for m in meetings],
        })

    if ambiental_applies:
        # Q-11 (spec Sec.1): associacao do Parecer (PA-NNN) por CONTEUDO --
        # ver extratores/parecer.py. `parecer` e o dict devolvido por
        # `parecer.avaliar()`; None so quando o ciclo ambiental nao aplica
        # (chamador nao computa nesse caso).
        p = parecer or {"status": "NOT_FOUND", "motivo": "nao avaliado"}
        entrada = {"id": "parecer_tecnico", "obrigatorio": False,
                  "mecanismo": p.get("mecanismo")}
        if p["status"] == "MATCH":
            doc = p["documento"]
            entrada.update({
                "presente": True, "pa_numero": doc["numero"], "path": doc["path"],
                "evidencia": p.get("evidencia"),
            })
        elif p["status"] == "HITL":
            entrada.update({
                "presente": "HITL",
                "candidatos": [{"pa_numero": c["numero"], "path": c["path"]} for c in p["candidatos"]],
                "motivo": p.get("motivo"),
            })
        else:  # NOT_FOUND
            entrada.update({"presente": False, "motivo": p.get("motivo")})
        anexos.append(entrada)
        anexos.append({
            "id": "acompanhamento_ambiental", "obrigatorio": False,
            "presente": bool(fa_ambiental), "path": fa_ambiental,
        })
        anexos.append({
            "id": "acompanhamento_sst", "obrigatorio": False,
            "presente": bool(fa_sst), "path": fa_sst,
        })

    if nc_no_periodo:
        anexos.append({
            "id": "nao_conformidade", "obrigatorio": False, "presente": True,
            "itens": nc_no_periodo,
        })

    anexos.append({
        "id": "diagrama_unifilar", "obrigatorio": True, "posicao": "penultimo",
        "presente": bool(unifilar_exists), "path": unifilar_path,
    })
    anexos.append({
        "id": "diario_de_obras", "obrigatorio": True, "posicao": "ultimo",
        "presente": bool(diario_exists), "path": diario_path,
    })

    for a in anexos:
        _assembly.classificar_anexo(a["id"])  # levanta se faltar classificacao
    return anexos
