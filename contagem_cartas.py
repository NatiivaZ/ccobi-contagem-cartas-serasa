# -*- coding: utf-8 -*-
"""Contagem de cartas a partir da planilha de autos.

A regra continua a mesma do processo atual: conta por dia, por documento e sem
repetir o número do auto.
"""

import json
import math
import re
from pathlib import Path
from datetime import datetime, timedelta

import pandas as pd


# O JSON fica junto do script porque normalmente é ajustado na própria pasta do projeto.
CONFIG_PATH = Path(__file__).parent / "config_colunas.json"
DEFAULT_CONFIG = {
    "numero_auto": "Número do auto",
    "data_inscricao": "Data de inscrição",
    "cpf_cnpj": "CPF/CNPJ",
}


def carregar_config():
    """Lê a configuração das colunas e cai no padrão se o arquivo não existir."""
    if CONFIG_PATH.exists():
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    return DEFAULT_CONFIG.copy()


def _ler_planilha(caminho_excel):
    """Lê a planilha e dá uma limpada básica nos nomes das colunas."""
    df = pd.read_excel(caminho_excel, engine="openpyxl")
    df.columns = df.columns.str.strip()
    return df


def _validar_colunas(df, colunas):
    """Confere se a planilha trouxe todas as colunas que o processo precisa."""
    for coluna in colunas:
        if coluna not in df.columns:
            raise ValueError(f'Coluna "{coluna}" não encontrada na planilha. Colunas: {list(df.columns)}')


def _formatar_data_saida(valor):
    """Padroniza a data para o formato que já vai para os relatórios."""
    if valor is None or (hasattr(valor, "__len__") and len(str(valor)) == 0):
        return ""
    if hasattr(valor, "strftime"):
        return valor.strftime("%d/%m/%Y")
    return str(valor)


def _montar_totais(total_cartas, total_autos, linhas_sem_data, linhas_invalidas, autos_duplicados_removidos):
    """Junta os totais em um único dicionário para não espalhar esse resumo."""
    return {
        "total_cartas": int(total_cartas),
        "total_autos": int(total_autos),
        "linhas_ignoradas_sem_data": int(linhas_sem_data),
        "linhas_data_invalida": int(linhas_invalidas),
        "autos_duplicados_removidos": int(autos_duplicados_removidos),
    }


def _emitir_resumo_console(out_path, totais, linhas_inv, por_dia):
    """Mostra no console o mesmo resumo que o usuário já espera na versão CLI."""
    if linhas_inv:
        print(
            f"Atenção: {len(linhas_inv)} linha(s) com data inválida (foram ignoradas). "
            f"Linhas: {linhas_inv[:20]}{'...' if len(linhas_inv) > 20 else ''}"
        )

    if por_dia is not None and not por_dia.empty:
        soma_cartas_dia = int(por_dia["cartas"].sum())
        if soma_cartas_dia != totais["total_cartas"]:
            print(f"Aviso: soma das cartas por dia ({soma_cartas_dia}) ≠ total geral ({totais['total_cartas']})")
        else:
            print("Verificação: soma por dia = total geral (OK)")

    print(f"Resultado salvo em: {out_path}")
    print(f"Total de cartas: {totais['total_cartas']} | Total de autos (únicos): {totais['total_autos']}")


def normalizar_cpf_cnpj(val):
    """Deixa o documento só com números para o agrupamento não depender de máscara."""
    if pd.isna(val):
        return None
    s = str(val).strip()
    if not s:
        return None
    digits = re.sub(r"\D", "", s)
    return digits if digits else None


def parsear_data(val):
    """Tenta transformar o valor em data.

    Aceita texto, `datetime` e também número serial do Excel. O retorno segue no
    formato `(data, data_invalida)` porque o processo usa as duas informações.
    """
    if pd.isna(val):
        return None, False
    if isinstance(val, datetime):
        return val.date(), False
    # Em algumas planilhas a data vem como número serial do Excel.
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
    # Se veio como texto, tenta os formatos mais comuns das planilhas daqui.
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
    return None, True


def extrair_datas_planilha(caminho_excel, config=None):
    """Lê a planilha e devolve as datas válidas usadas no filtro da interface."""
    if config is None:
        config = carregar_config()
    col_data = config["data_inscricao"]
    df = _ler_planilha(caminho_excel)
    _validar_colunas(df, [col_data])
    datas_parsed = df[col_data].map(parsear_data)
    datas_ok = [x[0] for x in datas_parsed if x[0] is not None]
    datas_unicas = sorted(set(datas_ok))
    meses_anos = sorted(set((d.month, d.year) for d in datas_unicas))
    return datas_unicas, meses_anos


def processar_planilha(caminho_excel, config=None, datas_selecionadas=None):
    """Processa a planilha inteira e devolve os DataFrames usados na exportação."""
    if config is None:
        config = carregar_config()
    col_auto = config["numero_auto"]
    col_data = config["data_inscricao"]
    col_cpf = config["cpf_cnpj"]

    df = _ler_planilha(caminho_excel)
    _validar_colunas(df, [col_auto, col_data, col_cpf])

    # O agrupamento trabalha com documento limpo para evitar diferença só de máscara.
    df["_cpf_cnpj_norm"] = df[col_cpf].map(normalizar_cpf_cnpj)

    # A data já sai em um formato pronto para filtro e agrupamento.
    datas_parsed = df[col_data].map(parsear_data)
    df["_data"] = [x[0] for x in datas_parsed]
    df["_data_invalida"] = [x[1] for x in datas_parsed]

    # Sem data não entra na conta.
    df_ok = df[df["_data"].notna()].copy()
    linhas_invalidas = df[df["_data_invalida"]].index.tolist()

    # Documento vazio ou inválido também não ajuda no agrupamento.
    df_ok = df_ok[df_ok["_cpf_cnpj_norm"].notna()]

    # Quando a GUI manda um recorte, filtra antes de deduplicar.
    if datas_selecionadas is not None:
        set_datas = set(datas_selecionadas)
        df_ok = df_ok[df_ok["_data"].isin(set_datas)]

    n_antes_dedup = len(df_ok)
    # O mesmo auto só pode contar uma vez.
    df_ok = df_ok.drop_duplicates(subset=[col_auto], keep="first")
    autos_duplicados_removidos = n_antes_dedup - len(df_ok)

    if df_ok.empty:
        totais_vazio = _montar_totais(
            total_cartas=0,
            total_autos=0,
            linhas_sem_data=len(df[df["_data"].isna()]),
            linhas_invalidas=len(linhas_invalidas),
            autos_duplicados_removidos=autos_duplicados_removidos,
        )
        return None, None, None, totais_vazio, linhas_invalidas

    # Nesse ponto os autos já estão limpos e únicos, então o agrupamento fica direto.
    agrupado = (
        df_ok.groupby(["_data", "_cpf_cnpj_norm"], as_index=False)
        .agg(autos=(col_auto, "count"))
        .rename(columns={"autos": "autos"})
    )
    # Regra da carta: cada bloco de até 10 autos vira 1 carta.
    agrupado["cartas"] = agrupado["autos"].apply(lambda n: math.ceil(n / 10))

    # Resumo por dia.
    por_dia = (
        agrupado.groupby("_data", as_index=False)
        .agg(cartas=("cartas", "sum"), autos=("autos", "sum"))
        .rename(columns={"_data": "data"})
    )
    por_dia["data"] = por_dia["data"].map(_formatar_data_saida)

    # Resumo por documento no período inteiro.
    por_cnpj = (
        agrupado.groupby("_cpf_cnpj_norm", as_index=False)
        .agg(cartas=("cartas", "sum"), autos=("autos", "sum"))
        .rename(columns={"_cpf_cnpj_norm": "cpf_cnpj"})
    )

    # Detalhe diário por documento.
    por_cnpj_dia = agrupado.copy()
    por_cnpj_dia["data"] = por_cnpj_dia["_data"].map(_formatar_data_saida)
    por_cnpj_dia = por_cnpj_dia.rename(columns={"_cpf_cnpj_norm": "cpf_cnpj"})[["data", "cpf_cnpj", "autos", "cartas"]]

    total_cartas = int(agrupado["cartas"].sum())
    total_autos = int(df_ok[col_auto].count())

    linhas_sem_data = (df["_data"].isna()).sum() if "_data" in df.columns else 0
    totais = _montar_totais(
        total_cartas=total_cartas,
        total_autos=total_autos,
        linhas_sem_data=linhas_sem_data,
        linhas_invalidas=len(linhas_invalidas),
        autos_duplicados_removidos=autos_duplicados_removidos,
    )

    return por_dia, por_cnpj, por_cnpj_dia, totais, linhas_invalidas


def exportar_resultados(caminho_excel, pasta_saida=None, config=None, datas_selecionadas=None):
    """Processa a planilha e grava o resultado final em um novo arquivo Excel."""
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
        # A aba Resumo vem primeiro porque costuma ser a mais consultada.
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

    _emitir_resumo_console(out_path, totais, linhas_inv, por_dia)
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
