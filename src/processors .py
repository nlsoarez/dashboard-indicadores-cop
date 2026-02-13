import pandas as pd
import numpy as np
import openpyxl
import io
from src.config import (
    EQUIPE_IDS, BASE_EQUIPE, HEADER_ROW, SHEET_NAME_CANDIDATES,
    COL_LOGIN, COL_NOME, COL_BASE, COL_DATA, COL_MES, COL_ANOMES,
    COL_VOL_TOTAL, COL_VOL_MEDIA, COL_DPA_USO, COL_DPA_JORNADA,
    COL_DPA_RESULTADO, VOL_COLS, COL_CARGO, COL_PERIODO, COL_COORD,
)


def list_sheets(uploaded_file):
    data = uploaded_file.getvalue()
    wb = openpyxl.load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    return wb.sheetnames


def load_produtividade(uploaded_file) -> pd.DataFrame:
    """Lê a planilha de produtividade e retorna DataFrame filtrado pela equipe."""
    sheets = list_sheets(uploaded_file)

    sheet_to_read = None
    for candidate in SHEET_NAME_CANDIDATES:
        if candidate in sheets:
            sheet_to_read = candidate
            break
    if sheet_to_read is None:
        sheet_to_read = sheets[0]

    df = pd.read_excel(uploaded_file, sheet_name=sheet_to_read, header=HEADER_ROW)

    # Remove coluna unnamed
    df = df.loc[:, ~df.columns.str.startswith("Unnamed")]

    # Filtra equipe
    df[COL_LOGIN] = df[COL_LOGIN].astype(str).str.strip()
    df_equipe = df[df[COL_LOGIN].isin(EQUIPE_IDS)].copy()

    # Garante tipos numéricos
    num_cols = [COL_VOL_TOTAL, COL_VOL_MEDIA, COL_DPA_USO, COL_DPA_JORNADA, COL_DPA_RESULTADO]
    num_cols += list(VOL_COLS.keys())
    for c in num_cols:
        if c in df_equipe.columns:
            df_equipe[c] = pd.to_numeric(df_equipe[c], errors="coerce")

    # Data como datetime
    if COL_DATA in df_equipe.columns:
        df_equipe[COL_DATA] = pd.to_datetime(df_equipe[COL_DATA], errors="coerce")

    # Merge com info da equipe (setor fixo)
    df_equipe = df_equipe.merge(
        BASE_EQUIPE[["Matricula", "Setor"]],
        left_on=COL_LOGIN, right_on="Matricula", how="left"
    )

    return df_equipe


def resumo_mensal(df: pd.DataFrame) -> pd.DataFrame:
    """Gera resumo mensal por analista."""
    agg_dict = {
        COL_DATA: "count",
        COL_VOL_TOTAL: "sum",
    }

    # Adiciona volumes individuais
    for vc in VOL_COLS.keys():
        if vc in df.columns:
            agg_dict[vc] = "sum"

    # DPA: média (filtrando outliers)
    group_cols = [COL_LOGIN, COL_NOME, "Setor", COL_MES, COL_ANOMES]
    existing_group = [c for c in group_cols if c in df.columns]

    g = df.groupby(existing_group).agg(agg_dict).reset_index()
    g = g.rename(columns={COL_DATA: "Dias_Trabalhados"})

    # Média diária
    g["Media_Diaria"] = (g[COL_VOL_TOTAL] / g["Dias_Trabalhados"]).round(1)

    # DPA média por analista/mês (calculada separadamente para filtrar outliers)
    dpa_valid = df[(df[COL_DPA_RESULTADO] >= 0) & (df[COL_DPA_RESULTADO] <= 120)].copy()
    if not dpa_valid.empty:
        dpa_mean = dpa_valid.groupby(existing_group)[COL_DPA_RESULTADO].mean().reset_index()
        dpa_mean = dpa_mean.rename(columns={COL_DPA_RESULTADO: "DPA_Media"})
        dpa_mean["DPA_Media"] = dpa_mean["DPA_Media"].round(1)
        g = g.merge(dpa_mean, on=existing_group, how="left")
    else:
        g["DPA_Media"] = np.nan

    return g


def resumo_geral(df: pd.DataFrame) -> pd.DataFrame:
    """Gera resumo geral (todos os meses) por analista."""
    agg_dict = {
        COL_DATA: "count",
        COL_VOL_TOTAL: "sum",
    }
    for vc in VOL_COLS.keys():
        if vc in df.columns:
            agg_dict[vc] = "sum"

    group_cols = [COL_LOGIN, COL_NOME, "Setor"]
    existing_group = [c for c in group_cols if c in df.columns]

    g = df.groupby(existing_group).agg(agg_dict).reset_index()
    g = g.rename(columns={COL_DATA: "Dias_Trabalhados"})
    g["Media_Diaria"] = (g[COL_VOL_TOTAL] / g["Dias_Trabalhados"]).round(1)

    # DPA
    dpa_valid = df[(df[COL_DPA_RESULTADO] >= 0) & (df[COL_DPA_RESULTADO] <= 120)].copy()
    if not dpa_valid.empty:
        dpa_mean = dpa_valid.groupby(existing_group)[COL_DPA_RESULTADO].mean().reset_index()
        dpa_mean = dpa_mean.rename(columns={COL_DPA_RESULTADO: "DPA_Media"})
        dpa_mean["DPA_Media"] = dpa_mean["DPA_Media"].round(1)
        g = g.merge(dpa_mean, on=existing_group, how="left")
    else:
        g["DPA_Media"] = np.nan

    return g


def evolucao_diaria(df: pd.DataFrame) -> pd.DataFrame:
    """Volume total diário da equipe inteira."""
    if COL_DATA not in df.columns:
        return pd.DataFrame()
    daily = df.groupby(COL_DATA).agg(
        Vol_Total=(COL_VOL_TOTAL, "sum"),
        Analistas=(COL_LOGIN, "nunique"),
    ).reset_index()
    daily["Media_por_Analista"] = (daily["Vol_Total"] / daily["Analistas"]).round(1)
    return daily


def composicao_volume(df: pd.DataFrame) -> pd.DataFrame:
    """Composição do volume por tipo de atividade."""
    vol_data = {}
    for col, label in VOL_COLS.items():
        if col in df.columns:
            total = df[col].sum()
            if total > 0:
                vol_data[label] = total
    return pd.DataFrame(list(vol_data.items()), columns=["Atividade", "Volume"]).sort_values(
        "Volume", ascending=False
    )


def primeiro_nome(nome_completo: str) -> str:
    """Retorna primeiro + último nome para display compacto."""
    parts = str(nome_completo).strip().split()
    if len(parts) <= 2:
        return nome_completo
    return f"{parts[0]} {parts[-1]}"
