#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
t13_clima_dia_sem_ficha.py

Descobre qual regra o redator usa para preencher o CLIMA de um dia
que não tem Ficha de Campo.

H1: valor default fixo ("Bom")
H2: repetição do clima do dia anterior

Estratégia:
1. Para cada relatório, extrair (data, clima) dos 7 dias
2. Para cada obra, listar datas com Ficha de Campo
3. Cruzar: achar dias sem ficha mas com clima declado
4. Achar casos discriminantes: dia anterior != "Bom"
5. Contar quantos apoiam H1, quantos apoiam H2
"""
import json
import re
import os
from datetime import datetime, timedelta
from pathlib import Path
from collections import defaultdict
import sys

# Adicionar scripts ao path para usar prog_util
sys.path.insert(0, os.path.dirname(__file__))
import prog_util

PROJETO = prog_util.PROJ
OUT = prog_util.OUT
TEXTO_DIR = os.path.join(OUT, "extracao", "texto")
INVENTARIO_FILE = os.path.join(OUT, "inventarios", "relatorios_semanais_historico.json")
ANALISES_DIR = os.path.join(OUT, "analises")
OUTPUT_FILE = os.path.join(ANALISES_DIR, "t13_clima_dia_sem_ficha.json")

os.makedirs(ANALISES_DIR, exist_ok=True)

# Mês português -> número
MESES_PT = {
    "jan": 1, "janeiro": 1,
    "fev": 2, "fevereiro": 2,
    "mar": 3, "março": 3,
    "abr": 4, "abril": 4,
    "mai": 5, "maio": 5,
    "jun": 6, "junho": 6,
    "jul": 7, "julho": 7,
    "ago": 8, "agosto": 8,
    "set": 9, "setembro": 9,
    "out": 10, "outubro": 10,
    "nov": 11, "novembro": 11,
    "dez": 12, "dezembro": 12,
}

DIA_SEMANA_PT = {
    "segunda": "monday", "segunda-feira": "monday",
    "terça": "tuesday", "terça-feira": "tuesday",
    "quarta": "wednesday", "quarta-feira": "wednesday",
    "quinta": "thursday", "quinta-feira": "thursday",
    "sexta": "friday", "sexta-feira": "friday",
    "sábado": "saturday",
    "domingo": "sunday",
}

def normalizar_texto(s):
    """Remove espaços extras, minúsculas."""
    return re.sub(r'\s+', ' ', str(s).strip().lower())

def extrair_tabela_clima(texto):
    """
    Extrai (data, clima) da tabela "DATA: ... DIA: ... CLIMA: ...".

    Retorna lista de tuplas (data_str, dia_semana_str, clima_str).
    Tolerante a quebras de linha.
    """
    # Normalizar espaços mas preservar estrutura
    texto_norm = re.sub(r'\s+', ' ', texto)

    # Procura por: DATA: ...valores... DIA: ...valores... CLIMA: ...valores...
    # Parar quando encontra: "4.X." (seções) ou linha que começa com maiúscula após espaço
    match = re.search(
        r'DATA:\s*(.*?)\s+DIA:\s*(.*?)\s+CLIMA:\s*(.*?)(?:\n[A-Z]{2,}|4\.[0-9]|$)',
        texto_norm,
        re.IGNORECASE | re.DOTALL
    )

    if not match:
        return []

    datas_str = match.group(1)
    dias_str = match.group(2)
    climas_str = match.group(3)

    # Split por números de data (DD/ ou DD.) para datas
    # Esperamos formato: 25/set 26/set ou 25/set. 26/set. ou 07/nov. 08/nov.
    datas = re.findall(r'(\d{1,2})[/\.]([a-z]{3})', datas_str, re.IGNORECASE)

    # Split dias: SEGUNDA TERÇA QUARTA ... (pode quebrar em linhas)
    # Remover pontos e normalizar
    dias_norm = normalizar_texto(dias_str)
    dias_parts = re.split(r'\s+', dias_norm)
    dias_parts = [d for d in dias_parts if d]

    # Reconstruir dias (alguns podem ser "segunda-feira" quebrado em "segunda feira")
    dias = []
    i = 0
    while i < len(dias_parts):
        if i + 1 < len(dias_parts) and dias_parts[i+1] == "feira":
            dias.append(dias_parts[i] + "-feira")
            i += 2
        else:
            dias.append(dias_parts[i])
            i += 1

    # Split climas: BOM CHUVOSO IMPRATICÁVEL ...
    # Validar que são valores conhecidos
    climas_norm = normalizar_texto(climas_str)
    climas_parts = re.split(r'\s+', climas_norm)

    # Filtrar apenas valores válidos de clima (até encontrar algo que não é clima)
    climas_validos = ["bom", "chuvoso", "impraticável", "impraticavel", "-"]
    climas = []
    for c in climas_parts:
        if c in climas_validos:
            climas.append(c)
        elif c:
            # Parou de ser clima, quebra
            break

    # Montar resultado
    resultado = []
    for j in range(len(datas)):
        if j < len(dias) and j < len(climas):
            dia_str = datas[j][0]
            mes_str = datas[j][1].lower()
            mes_num = MESES_PT.get(mes_str, 1)

            dia_semana = dias[j] if j < len(dias) else "unknown"
            clima = climas[j] if j < len(climas) else "unknown"

            resultado.append({
                "dia": dia_str,
                "mes": mes_num,
                "mes_abrev": mes_str,
                "dia_semana": dia_semana,
                "clima": clima
            })

    return resultado

def meses_portugues_para_numero(abrev):
    """Converte 'set', 'out', 'nov', etc para número."""
    return MESES_PT.get(abrev.lower(), None)

def listar_fichas_de_campo(obra):
    """
    Lista as datas que têm Ficha de Campo para uma obra.
    Retorna set de datas no formato YYYY-MM-DD.

    Fichas estão em: {BASE_OBRAS}\\{obra}\\3_Gestao Tecnica de Obra\\4. Relat*\\1. Ficha*\\
    Nomes: AAAA-MM-DD Ficha de Campo ....xlsx
    """
    gto = os.path.join(prog_util.BASE_OBRAS, obra, "3_Gestao Tecnica de Obra")

    # Encontrar pasta "4. Relat*"
    try:
        dirs_gto = prog_util.listdir(gto)
    except Exception as e:
        print(f"Erro listando {gto}: {e}", file=sys.stderr)
        return set()

    pasta_relat = None
    for d in dirs_gto:
        if d.startswith("4. Relat"):
            pasta_relat = os.path.join(gto, d)
            break

    if not pasta_relat:
        print(f"Pasta '4. Relat*' não encontrada para {obra}", file=sys.stderr)
        return set()

    # Encontrar pasta "1. Fichas de campo"
    try:
        dirs_relat = prog_util.listdir(pasta_relat)
    except Exception as e:
        print(f"Erro listando {pasta_relat}: {e}", file=sys.stderr)
        return set()

    pasta_fichas = None
    for d in dirs_relat:
        if d.startswith("1. Ficha"):
            pasta_fichas = os.path.join(pasta_relat, d)
            break

    if not pasta_fichas:
        print(f"Pasta '1. Ficha*' não encontrada para {obra}", file=sys.stderr)
        return set()

    # Listar arquivos: AAAA-MM-DD Ficha de Campo *.xlsx
    datas_fichas = set()
    try:
        arquivos = prog_util.listdir(pasta_fichas)
    except Exception as e:
        print(f"Erro listando {pasta_fichas}: {e}", file=sys.stderr)
        return set()

    for arq in arquivos:
        # Procura por padrão: AAAA-MM-DD Ficha de Campo
        match = re.match(r'(\d{4})-(\d{2})-(\d{2})\s+Ficha', arq, re.IGNORECASE)
        if match:
            data_str = f"{match.group(1)}-{match.group(2)}-{match.group(3)}"
            datas_fichas.add(data_str)

    return datas_fichas

def inferir_ano(dia, mes, periodo_inicio, periodo_fim):
    """
    Infere o ano baseado no período do relatório.
    periodo_inicio e periodo_fim são datas (YYYY-MM-DD).
    """
    inicio = datetime.strptime(periodo_inicio, "%Y-%m-%d")
    fim = datetime.strptime(periodo_fim, "%Y-%m-%d")

    # Tentar criar data com ano do período
    try:
        data_tentativa = datetime(inicio.year, mes, int(dia))
        if data_tentativa >= inicio and data_tentativa <= fim:
            return inicio.year
    except ValueError:
        pass

    # Se não encaixou, tentar ano anterior
    try:
        data_tentativa = datetime(inicio.year - 1, mes, int(dia))
        if data_tentativa >= inicio and data_tentativa <= fim:
            return inicio.year - 1
    except ValueError:
        pass

    # Se não encaixou, tentar ano seguinte
    try:
        data_tentativa = datetime(inicio.year + 1, mes, int(dia))
        if data_tentativa >= inicio and data_tentativa <= fim:
            return inicio.year + 1
    except ValueError:
        pass

    # Default: ano do início
    return inicio.year

def processar_relatorio(relatorio_meta, texto_file):
    """
    Processa um relatório: extrai clima, identifica dias sem ficha.

    Retorna dicionário com resultado.
    """
    obra = relatorio_meta["obra"]
    numero = relatorio_meta["numero"]
    revisao = relatorio_meta.get("revisao", 1)
    periodo_inicio = relatorio_meta["pasta_data_inicio"]
    periodo_fim = relatorio_meta["pasta_data_fim"]

    # Ler texto
    try:
        with open(texto_file, 'r', encoding='utf-8') as f:
            texto = f.read()
    except Exception as e:
        return {
            "obra": obra,
            "numero": numero,
            "revisao": revisao,
            "status": "erro_leitura",
            "erro": str(e),
            "climas": [],
            "dias_sem_ficha": []
        }

    # Extrair tabela de clima
    tabela_clima = extrair_tabela_clima(texto)

    if not tabela_clima:
        return {
            "obra": obra,
            "numero": numero,
            "revisao": revisao,
            "status": "sem_tabela",
            "climas": [],
            "dias_sem_ficha": []
        }

    # Converter para datas reais
    climas_processados = []
    for item in tabela_clima:
        try:
            ano = inferir_ano(item["dia"], item["mes"], periodo_inicio, periodo_fim)
            data = datetime(ano, item["mes"], int(item["dia"]))
            data_str = data.strftime("%Y-%m-%d")
            climas_processados.append({
                "data": data_str,
                "dia_semana": item["dia_semana"],
                "clima": item["clima"]
            })
        except Exception as e:
            # Ignorar dias que não conseguir processar
            continue

    # Listar fichas de campo dessa obra
    fichas_datas = listar_fichas_de_campo(obra)

    # Encontrar dias sem ficha
    dias_sem_ficha = []
    for i, clima_item in enumerate(climas_processados):
        data = clima_item["data"]
        clima = clima_item["clima"]

        if data not in fichas_datas:
            # Dia sem ficha
            dias_sem_ficha.append({
                "indice": i,
                "data": data,
                "clima": clima,
                "dia_anterior": None,
                "clima_dia_anterior": None
            })

    # Preencher clima do dia anterior
    for item in dias_sem_ficha:
        idx = item["indice"]
        if idx > 0:
            item["dia_anterior"] = climas_processados[idx - 1]["data"]
            item["clima_dia_anterior"] = climas_processados[idx - 1]["clima"]

    return {
        "obra": obra,
        "numero": numero,
        "revisao": revisao,
        "periodo_inicio": periodo_inicio,
        "periodo_fim": periodo_fim,
        "status": "ok",
        "climas": climas_processados,
        "dias_sem_ficha": dias_sem_ficha
    }

def main():
    """Análise principal."""
    print("Carregando inventário...", file=sys.stderr)
    with open(INVENTARIO_FILE, 'r', encoding='utf-8') as f:
        inventario = json.load(f)

    relatorios = inventario.get("relatorios", [])
    print(f"Total de {len(relatorios)} relatórios", file=sys.stderr)

    resultados = []
    erros = []

    for i, rel_meta in enumerate(relatorios):
        print(f"[{i+1}/{len(relatorios)}] Processando {rel_meta['obra']} RO-{rel_meta['numero']}.{rel_meta.get('revisao', 1)}...", file=sys.stderr)

        # Arquivo de texto
        texto_cache = rel_meta.get("texto_cache")
        if not texto_cache:
            print(f"  Aviso: sem texto_cache", file=sys.stderr)
            continue

        # Converter caminho relativo para absoluto
        if not os.path.isabs(texto_cache):
            texto_file = os.path.join(PROJETO, texto_cache)
        else:
            texto_file = texto_cache

        if not os.path.exists(texto_file):
            print(f"  Erro: arquivo não existe: {texto_file}", file=sys.stderr)
            erros.append({
                "obra": rel_meta["obra"],
                "numero": rel_meta["numero"],
                "erro": "arquivo_nao_existe"
            })
            continue

        # Processar
        resultado = processar_relatorio(rel_meta, texto_file)
        resultados.append(resultado)

        if resultado["status"] != "ok":
            print(f"  Status: {resultado['status']}", file=sys.stderr)
        else:
            n_sem_ficha = len(resultado["dias_sem_ficha"])
            print(f"  OK: {n_sem_ficha} dias sem ficha", file=sys.stderr)

    # Estatísticas
    print("\n=== ESTATÍSTICAS ===", file=sys.stderr)

    sucesso = sum(1 for r in resultados if r["status"] == "ok")
    print(f"Relatórios processados com sucesso: {sucesso}/{len(relatorios)}", file=sys.stderr)

    # Dias sem ficha por obra
    dias_sem_ficha_por_obra = defaultdict(list)
    for rel in resultados:
        if rel["status"] == "ok":
            for dia in rel["dias_sem_ficha"]:
                dias_sem_ficha_por_obra[rel["obra"]].append({
                    "numero": rel["numero"],
                    "data": dia["data"],
                    "clima": dia["clima"],
                    "clima_anterior": dia["clima_dia_anterior"]
                })

    print("\nDias sem ficha encontrados:", file=sys.stderr)
    for obra in sorted(dias_sem_ficha_por_obra.keys()):
        n = len(dias_sem_ficha_por_obra[obra])
        print(f"  {obra}: {n} dias", file=sys.stderr)

    # Casos discriminantes
    casos_discriminantes = []
    h1_count = 0  # clima = "Bom"
    h2_count = 0  # clima = clima anterior

    for rel in resultados:
        if rel["status"] == "ok":
            for dia in rel["dias_sem_ficha"]:
                clima_anterior = dia["clima_dia_anterior"]

                # Discriminante: dia anterior != "Bom"
                if clima_anterior and normalizar_texto(clima_anterior) != "bom":
                    clima = dia["clima"]
                    casos_discriminantes.append({
                        "obra": rel["obra"],
                        "numero": rel["numero"],
                        "data": dia["data"],
                        "dia_anterior": dia["dia_anterior"],
                        "clima_anterior": clima_anterior,
                        "clima_declarado": clima
                    })

                    # Contar apoio a H1 ou H2
                    if normalizar_texto(clima) == "bom":
                        h1_count += 1
                    elif normalizar_texto(clima) == normalizar_texto(clima_anterior):
                        h2_count += 1

    print(f"\nCasos discriminantes (dia anterior != 'Bom'): {len(casos_discriminantes)}", file=sys.stderr)
    print(f"  Apoiam H1 (clima = 'Bom'): {h1_count}", file=sys.stderr)
    print(f"  Apoiam H2 (clima = clima anterior): {h2_count}", file=sys.stderr)

    # Veredito
    if len(casos_discriminantes) == 0:
        veredito = "INCONCLUSIVO"
        razao = "Nenhum caso discriminante encontrado (todos têm dia anterior = 'Bom')"
    elif h1_count > 0 and h2_count == 0:
        veredito = "H1"
        razao = f"Todos os {h1_count} casos discriminantes declaram clima='Bom'"
    elif h2_count > 0 and h1_count == 0:
        veredito = "H2"
        razao = f"Todos os {h2_count} casos discriminantes repetem clima do dia anterior"
    elif h1_count > h2_count:
        veredito = "H1_PREDOMINANTE"
        razao = f"Maioria ({h1_count}/{h1_count+h2_count}) declara 'Bom'"
    elif h2_count > h1_count:
        veredito = "H2_PREDOMINANTE"
        razao = f"Maioria ({h2_count}/{h1_count+h2_count}) repete clima anterior"
    else:
        veredito = "INCONCLUSIVO"
        razao = "Número igual de casos para H1 e H2"

    print(f"\nVEREDITO: {veredito}", file=sys.stderr)
    print(f"Razão: {razao}", file=sys.stderr)

    # Preparar saída
    output_data = {
        "gerado_em": datetime.now().isoformat(),
        "gerado_por": __file__,
        "resumo": {
            "total_relatorios": len(relatorios),
            "relatorios_sucesso": sucesso,
            "dias_sem_ficha_total": sum(len(dias) for dias in dias_sem_ficha_por_obra.values()),
            "dias_sem_ficha_por_obra": {
                obra: len(dias)
                for obra, dias in sorted(dias_sem_ficha_por_obra.items())
            },
            "casos_discriminantes": len(casos_discriminantes),
            "h1_count": h1_count,
            "h2_count": h2_count,
            "veredito": veredito,
            "razao": razao
        },
        "casos_discriminantes": sorted(casos_discriminantes, key=lambda x: (x["obra"], x["data"])),
        "relatorios": resultados
    }

    # Salvar
    os.makedirs(ANALISES_DIR, exist_ok=True)
    with open(OUTPUT_FILE, 'w', encoding='utf-8') as f:
        json.dump(output_data, f, indent=2, ensure_ascii=False)

    print(f"\nResultado salvo em: {OUTPUT_FILE}", file=sys.stderr)

if __name__ == "__main__":
    main()
