# -*- coding: utf-8 -*-
"""scripts/gerador/renderers/ - motores de renderizacao do document_model.

Cada renderer consome so `gerador.document_model.Document` (nunca report_data
nem caminho de planilha) e produz um formato de saida. Hoje so existe
`docx_renderer`; adicionar outro formato nao exige mexer em content_builder.
"""
