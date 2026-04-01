# -*- coding: utf-8 -*-
"""
Automação: Contagem de Cartas a partir de planilha de autos de infração.
Regra: 1-10 autos (mesmo CPF/CNPJ, mesmo dia) = 1 carta; 11-20 = 2 cartas; etc.
Contagem por dia (não soma entre dias). Sem duplicar Número do auto.
"""

import json
import math
import re
from pathlib import Path
from datetime import datetime, timedelta

import pandas as pd


# Caminho do arquivo de configuração das colunas
CONFIG_PATH = Path(__file__).parent / "config_colunas.json"


def carregar_config():
    """Carrega nomes das colunas do JSON. Se não existir, usa padrão."""
    if CONFIG_PATH.exists():
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    return {
        "numero_auto": "Número do auto",
        "data_inscricao": "Data de inscrição",
        "cpf_cnpj": "CPF/CNPJ",
    }


def normalizar_cpf_cnpj(val):
    """Comparação por apenas dígitos (ignora pontos, traços, barras, espaços)."""
    if pd.isna(val):
        return None
    s = str(val).strip()
    if not s:
        return None
    digits = re.sub(r"\D", "", s)
    return digits if digits else None


def parsear_data(val):
    """Converte para data. Aceita dd/mm/aaaa, datetime ou número serial do Excel. Retorna (date ou None, inválido: bool)."""
    if pd.isna(val):
        return None, False
    if isinstance(val, datetime):
        return val.date(), False
    # Excel às vezes lê data como número serial (dias desde 1899-12-30); só aceita faixa plausível
    if isinstance(val, (int, float)):
        v = int(val)
        if 1 <= v <= 2958465:  # 1 = 1900-01-01, 2958465 ≈ ano 9999
            try:
                base = datetime(1899, 12, 30).date()
                return (base + timedelta(days=v)), False
            except (ValueError, OverflowError):
                pass
        return None, True
    s = str(val).strip()
    if not s:
        return None, False
    # dd/mm/aaaa ou dd-mm-aaaa
    for sep in ["/", "-"]:
        parts = s.split(sep)
        if len(parts) == 3:
            try:
                d, m, a = int(parts[0]), int(parts[1]), int(parts[2])
                if a < 100:
                    a += 2000 if a < 50 else 1900
                dt = datetime(a, m, d).date()
                return dt, False
            except (ValueError, TypeError):
                pass
    return None, True  # inválido


def extrair_datas_planilha(caminho_excel, config=None):
    """
    Lê a planilha e retorna as datas únicas (válidas) da coluna de data de inscrição,
    ordenadas. Útil para a GUI exibir filtro por mês ou por datas específicas.
    Retorna: (lista de date, lista de (mes, ano) únicos para dropdown de mês).
    """
    if config is None:
        config = carregar_config()
    col_data = config["data_inscricao"]
    df = pd.read_excel(caminho_excel, engine="openpyxl")
    df.columns = df.columns.str.strip()
    if col_data not in df.columns:
        raise ValueError(f'Coluna "{col_data}" não encontrada. Colunas: {list(df.columns)}')
    datas_parsed = df[col_data].map(parsear_data)
    datas_ok = [x[0] for x in datas_parsed if x[0] is not None]
    datas_unicas = sorted(set(datas_ok))
    meses_anos = sorted(set((d.month, d.year) for d in datas_unicas))
    return datas_unicas, meses_anos


def processar_planilha(caminho_excel, config=None, datas_selecionadas=None):
    """
    Lê o Excel, limpa dados, remove autos duplicados por número,
    agrupa por (data, CPF/CNPJ) e calcula cartas por dia.
    datas_selecionadas: set ou lista de date para filtrar (só processa essas datas); None = todas.
    Retorna (df_por_dia, df_por_cnpj, df_por_cnpj_dia, totais, linhas_invalidas).
    """
    if config is None:
        config = carregar_config()
    col_auto = config["numero_auto"]
    col_data = config["data_inscricao"]
    col_cpf = config["cpf_cnpj"]

    df = pd.read_excel(caminho_excel, engine="openpyxl")
    # Garantir nomes de coluna como estão no arquivo (strip)
    df.columns = df.columns.str.strip()

    for c in [col_auto, col_data, col_cpf]:
        if c not in df.columns:
            raise ValueError(f'Coluna "{c}" não encontrada na planilha. Colunas: {list(df.columns)}')

    # Normalizar CPF/CNPJ
    df["_cpf_cnpj_norm"] = df[col_cpf].map(normalizar_cpf_cnpj)
    # Datas
    datas_parsed = df[col_data].map(parsear_data)
    df["_data"] = [x[0] for x in datas_parsed]
    df["_data_invalida"] = [x[1] for x in datas_parsed]

    # Ignorar linhas sem data
    df_ok = df[df["_data"].notna()].copy()
    linhas_invalidas = df[df["_data_invalida"]].index.tolist()

    # Ignorar linhas sem CPF/CNPJ válido
    df_ok = df_ok[df_ok["_cpf_cnpj_norm"].notna()]

    # Filtro por datas (mês ou datas específicas)
    if datas_selecionadas is not None:
        set_datas = set(datas_selecionadas)
        df_ok = df_ok[df_ok["_data"].isin(set_datas)]

    n_antes_dedup = len(df_ok)
    # Remover duplicados por Número do auto (manter primeira ocorrência)
    df_ok = df_ok.drop_duplicates(subset=[col_auto], keep="first")
    autos_duplicados_removidos = n_antes_dedup - len(df_ok)

    if df_ok.empty:
        totais_vazio = {
            "total_cartas": 0,
            "total_autos": 0,
            "linhas_ignoradas_sem_data": len(df[df["_data"].isna()]),
            "linhas_data_invalida": len(linhas_invalidas),
            "autos_duplicados_removidos": autos_duplicados_removidos,
        }
        return None, None, None, totais_vazio, linhas_invalidas

    # Agrupar por (data, cpf_cnpj) e contar autos (já únicos por número do auto)
    agrupado = (
        df_ok.groupby(["_data", "_cpf_cnpj_norm"], as_index=False)
        .agg(autos=(col_auto, "count"))
        .rename(columns={"autos": "autos"})
    )
    # Cartas: teto(autos/10)
    agrupado["cartas"] = agrupado["autos"].apply(lambda n: math.ceil(n / 10))

    # --- Total por dia ---
    por_dia = (
        agrupado.groupby("_data", as_index=False)
        .agg(cartas=("cartas", "sum"), autos=("autos", "sum"))
        .rename(columns={"_data": "data"})
    )
    # Formato dd/mm/aaaa na saída (padrão combinado)
    def fmt_data(d):
        if d is None or (hasattr(d, '__len__') and len(str(d)) == 0):
            return ""
        if hasattr(d, 'strftime'):
            return d.strftime("%d/%m/%Y")
        return str(d)
    por_dia["data"] = por_dia["data"].map(fmt_data)

    # --- Total por CNPJ (todo o período) ---
    por_cnpj = (
        agrupado.groupby("_cpf_cnpj_norm", as_index=False)
        .agg(cartas=("cartas", "sum"), autos=("autos", "sum"))
        .rename(columns={"_cpf_cnpj_norm": "cpf_cnpj"})
    )

    # --- Por CNPJ por dia ---
    por_cnpj_dia = agrupado.copy()
    por_cnpj_dia["data"] = por_cnpj_dia["_data"].map(fmt_data)
    por_cnpj_dia = por_cnpj_dia.rename(columns={"_cpf_cnpj_norm": "cpf_cnpj"})[["data", "cpf_cnpj", "autos", "cartas"]]

    total_cartas = int(agrupado["cartas"].sum())
    total_autos = int(df_ok[col_auto].count())

    linhas_sem_data = (df["_data"].isna()).sum() if "_data" in df.columns else 0
    totais = {
        "total_cartas": total_cartas,
        "total_autos": total_autos,
        "linhas_ignoradas_sem_data": int(linhas_sem_data),
        "linhas_data_invalida": len(linhas_invalidas),
        "autos_duplicados_removidos": autos_duplicados_removidos,
    }

    return por_dia, por_cnpj, por_cnpj_dia, totais, linhas_invalidas


def exportar_resultados(caminho_excel, pasta_saida=None, config=None, datas_selecionadas=None):
    """
    Processa o Excel e grava resultados em uma nova planilha no mesmo arquivo
    (ou em arquivo de saída) com abas: Resumo, Por dia, Por CNPJ, Por CNPJ por dia.
    datas_selecionadas: set ou lista de date para filtrar; None = processar todas as datas.
    """
    caminho = Path(caminho_excel)
    if not caminho.exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {caminho_excel}")

    por_dia, por_cnpj, por_cnpj_dia, totais, linhas_inv = processar_planilha(
        caminho_excel, config, datas_selecionadas=datas_selecionadas
    )

    if pasta_saida:
        out_path = Path(pasta_saida) / (caminho.stem + "_resultado_cartas.xlsx")
    else:
        out_path = caminho.parent / (caminho.stem + "_resultado_cartas.xlsx")

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        # Resumo
        resumo = pd.DataFrame([
            ["Total de cartas", totais["total_cartas"]],
            ["Total de autos (únicos)", totais["total_autos"]],
            ["Linhas ignoradas (sem data ou data inválida)", totais.get("linhas_ignoradas_sem_data", 0)],
            ["Dessas, linhas com data inválida (formato incorreto)", totais.get("linhas_data_invalida", 0)],
            ["Autos duplicados (removidos)", totais.get("autos_duplicados_removidos", 0)],
        ])
        resumo.to_excel(writer, sheet_name="Resumo", index=False, header=False)

        if por_dia is not None:
            por_dia.to_excel(writer, sheet_name="Por dia", index=False)
            por_cnpj.to_excel(writer, sheet_name="Por CNPJ", index=False)
            por_cnpj_dia.to_excel(writer, sheet_name="Por CNPJ por dia", index=False)

    if linhas_inv:
        print(f"Atenção: {len(linhas_inv)} linha(s) com data inválida (foram ignoradas). Linhas: {linhas_inv[:20]}{'...' if len(linhas_inv) > 20 else ''}")

    # Verificação: soma por dia deve bater com total geral
    if por_dia is not None and not por_dia.empty:
        soma_cartas_dia = int(por_dia["cartas"].sum())
        if soma_cartas_dia != totais["total_cartas"]:
            print(f"Aviso: soma das cartas por dia ({soma_cartas_dia}) ≠ total geral ({totais['total_cartas']})")
        else:
            print("Verificação: soma por dia = total geral (OK)")

    print(f"Resultado salvo em: {out_path}")
    print(f"Total de cartas: {totais['total_cartas']} | Total de autos (únicos): {totais['total_autos']}")
    return out_path, totais


if __name__ == "__main__":
    import sys
    config = carregar_config()
    if len(sys.argv) < 2:
        print("Uso: python contagem_cartas.py <caminho_planilha.xlsx> [pasta_saida]")
        print("Colunas configuradas em config_colunas.json:", config)
        sys.exit(1)
    arquivo = sys.argv[1]
    pasta_saida = sys.argv[2] if len(sys.argv) > 2 else None
    exportar_resultados(arquivo, pasta_saida, config)
