import pandas as pd
import numpy as np
import openpyxl
import io
from src.config import (
    EQUIPE_IDS, BASE_EQUIPE, HEADER_ROW, SHEET_NAME_CANDIDATES,
    COL_LOGIN, COL_NOME, COL_BASE, COL_DATA, COL_MES, COL_ANOMES,
    COL_VOL_TOTAL, COL_VOL_MEDIA, COL_DPA_USO, COL_DPA_JORNADA,
    COL_DPA_RESULTADO, VOL_COLS, COL_CARGO, COL_PERIODO, COL_COORD,
    # ETIT
    ETIT_INDICADOR_FILTRO, ETIT_COL_INDICADOR, ETIT_COL_LOGIN,
    ETIT_COL_DEMANDA, ETIT_COL_VOLUME, ETIT_COL_STATUS,
    ETIT_COL_TIPO, ETIT_COL_AREA, ETIT_COL_CAUSA,
    ETIT_COL_REGIONAL, ETIT_COL_GRUPO, ETIT_COL_CIDADE, ETIT_COL_UF,
    ETIT_COL_TOA, ETIT_COL_DT_INICIO, ETIT_COL_DT_FIM,
    ETIT_COL_DT_ACIONAMENTO, ETIT_COL_TURNO,
    ETIT_COL_TMA, ETIT_COL_TMR, ETIT_COL_ANOMES,
    ETIT_SHEET_CANDIDATES, ETIT_COL_INDICADOR_VAL,
    # Residencial Indicadores
    RES_INDICADORES_FILTRO, RES_IND_INVERTIDOS, RES_SHEET_CANDIDATES,
    RES_COL_INDICADOR_NOME, RES_COL_ID_MOSTRA, RES_COL_VOLUME,
    RES_COL_INDICADOR_VAL, RES_COL_STATUS, RES_COL_REGIONAL,
    RES_COL_GRUPO, RES_COL_CIDADE, RES_COL_UF, RES_COL_TECNOLOGIA,
    RES_COL_SERVICO, RES_COL_NATUREZA, RES_COL_SINTOMA,
    RES_COL_FERRAMENTA, RES_COL_FECHAMENTO, RES_COL_SOLUCAO,
    RES_COL_IMPACTO, RES_COL_ENVIADO_TOA, RES_COL_DT_INICIO,
    RES_COL_DT_FIM, RES_COL_TMA, RES_COL_TMR, RES_COL_ANOMES,
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


# =====================================================
# ETIT POR EVENTO — Loader e processadores
# =====================================================
def load_etit(uploaded_file) -> pd.DataFrame:
    """Lê a planilha Analítico Empresarial e retorna apenas ETIT POR EVENTO da equipe."""
    sheets = list_sheets(uploaded_file)
    if hasattr(uploaded_file, 'seek'):
        uploaded_file.seek(0)

    sheet_to_read = None
    for candidate in ETIT_SHEET_CANDIDATES:
        if candidate in sheets:
            sheet_to_read = candidate
            break
    if sheet_to_read is None:
        sheet_to_read = sheets[0]

    df = pd.read_excel(uploaded_file, sheet_name=sheet_to_read)

    # Filtra apenas ETIT POR EVENTO
    if ETIT_COL_INDICADOR in df.columns:
        df = df[df[ETIT_COL_INDICADOR] == ETIT_INDICADOR_FILTRO].copy()
    else:
        return pd.DataFrame()

    if df.empty:
        return df

    # Filtra equipe
    df[ETIT_COL_LOGIN] = df[ETIT_COL_LOGIN].astype(str).str.strip()
    df_equipe = df[df[ETIT_COL_LOGIN].isin(EQUIPE_IDS)].copy()

    # Merge com info da equipe (nome e setor)
    df_equipe = df_equipe.merge(
        BASE_EQUIPE[["Matricula", "Nome", "Setor"]],
        left_on=ETIT_COL_LOGIN, right_on="Matricula", how="left"
    )

    # Garante tipos
    for c in [ETIT_COL_TMA, ETIT_COL_TMR, ETIT_COL_VOLUME, ETIT_COL_INDICADOR_VAL]:
        if c in df_equipe.columns:
            df_equipe[c] = pd.to_numeric(df_equipe[c], errors="coerce")

    for c in [ETIT_COL_DT_INICIO, ETIT_COL_DT_FIM, ETIT_COL_DT_ACIONAMENTO]:
        if c in df_equipe.columns:
            df_equipe[c] = pd.to_datetime(df_equipe[c], errors="coerce")

    if ETIT_COL_ANOMES in df_equipe.columns:
        df_equipe[ETIT_COL_ANOMES] = df_equipe[ETIT_COL_ANOMES].astype(str).str.strip()

    return df_equipe


def etit_resumo_analista(df: pd.DataFrame) -> pd.DataFrame:
    """Resumo ETIT por analista: total eventos, aderência, TMA/TMR médios, por demanda."""
    if df.empty:
        return pd.DataFrame()

    group = df.groupby([ETIT_COL_LOGIN, "Nome", "Setor"]).agg(
        Total_Eventos=(ETIT_COL_VOLUME, "sum"),
        Eventos_Aderentes=(ETIT_COL_INDICADOR_VAL, "sum"),
        TMA_Medio=(ETIT_COL_TMA, "mean"),
        TMR_Medio=(ETIT_COL_TMR, "mean"),
        RAL_Count=(ETIT_COL_DEMANDA, lambda x: (x == "RAL").sum()),
        REC_Count=(ETIT_COL_DEMANDA, lambda x: (x == "REC").sum()),
    ).reset_index()

    group["Aderencia_Pct"] = (group["Eventos_Aderentes"] / group["Total_Eventos"] * 100).round(1)
    group["TMA_Medio"] = group["TMA_Medio"].round(4)
    group["TMR_Medio"] = group["TMR_Medio"].round(4)
    group = group.sort_values("Total_Eventos", ascending=False).reset_index(drop=True)

    return group


def etit_por_demanda(df: pd.DataFrame) -> pd.DataFrame:
    """Breakdown por tipo de demanda (RAL/REC)."""
    if df.empty:
        return pd.DataFrame()
    return df.groupby(ETIT_COL_DEMANDA).agg(
        Eventos=(ETIT_COL_VOLUME, "sum"),
        Aderentes=(ETIT_COL_INDICADOR_VAL, "sum"),
        TMA_Medio=(ETIT_COL_TMA, "mean"),
        TMR_Medio=(ETIT_COL_TMR, "mean"),
    ).reset_index().rename(columns={ETIT_COL_DEMANDA: "Demanda"})


def etit_por_tipo(df: pd.DataFrame) -> pd.DataFrame:
    """Breakdown por tipo de rede/equipamento."""
    if df.empty:
        return pd.DataFrame()
    return df.groupby(ETIT_COL_TIPO).agg(
        Eventos=(ETIT_COL_VOLUME, "sum"),
        Aderentes=(ETIT_COL_INDICADOR_VAL, "sum"),
    ).reset_index().rename(columns={ETIT_COL_TIPO: "Tipo"}).sort_values("Eventos", ascending=False)


def etit_por_causa(df: pd.DataFrame) -> pd.DataFrame:
    """Breakdown por causa."""
    if df.empty:
        return pd.DataFrame()
    return df.groupby(ETIT_COL_CAUSA).agg(
        Eventos=(ETIT_COL_VOLUME, "sum"),
        Aderentes=(ETIT_COL_INDICADOR_VAL, "sum"),
    ).reset_index().rename(columns={ETIT_COL_CAUSA: "Causa"}).sort_values("Eventos", ascending=False)


def etit_por_regional(df: pd.DataFrame) -> pd.DataFrame:
    """Breakdown por regional."""
    if df.empty:
        return pd.DataFrame()
    return df.groupby(ETIT_COL_REGIONAL).agg(
        Eventos=(ETIT_COL_VOLUME, "sum"),
        Aderentes=(ETIT_COL_INDICADOR_VAL, "sum"),
    ).reset_index().rename(columns={ETIT_COL_REGIONAL: "Regional"}).sort_values("Eventos", ascending=False)


def etit_por_turno(df: pd.DataFrame) -> pd.DataFrame:
    """Breakdown por turno."""
    if df.empty:
        return pd.DataFrame()
    return df.groupby(ETIT_COL_TURNO).agg(
        Eventos=(ETIT_COL_VOLUME, "sum"),
        Aderentes=(ETIT_COL_INDICADOR_VAL, "sum"),
    ).reset_index().rename(columns={ETIT_COL_TURNO: "Turno"}).sort_values("Eventos", ascending=False)


def etit_evolucao_diaria(df: pd.DataFrame) -> pd.DataFrame:
    """Evolução diária de eventos ETIT."""
    if df.empty or ETIT_COL_DT_ACIONAMENTO not in df.columns:
        return pd.DataFrame()
    df_c = df.copy()
    df_c["Data"] = df_c[ETIT_COL_DT_ACIONAMENTO].dt.date
    daily = df_c.groupby("Data").agg(
        Eventos=(ETIT_COL_VOLUME, "sum"),
        Aderentes=(ETIT_COL_INDICADOR_VAL, "sum"),
        Analistas=(ETIT_COL_LOGIN, "nunique"),
    ).reset_index()
    daily["Data"] = pd.to_datetime(daily["Data"])
    daily["Aderencia_Pct"] = (daily["Aderentes"] / daily["Eventos"] * 100).round(1)
    return daily


# =====================================================
# INDICADORES RESIDENCIAL — Loader e processadores
# =====================================================

def load_residencial_indicadores(uploaded_file) -> pd.DataFrame:
    """
    Lê a planilha Analítico Indicadores Residencial e retorna apenas
    os 5 indicadores configurados, com tipagem correta.

    Nota: esta planilha NÃO possui coluna de login individual — os dados
    são agregados por evento/ocorrência (ID_MOSTRA).
    """
    sheets = list_sheets(uploaded_file)
    if hasattr(uploaded_file, "seek"):
        uploaded_file.seek(0)

    sheet_to_read = None
    for candidate in RES_SHEET_CANDIDATES:
        if candidate in sheets:
            sheet_to_read = candidate
            break
    if sheet_to_read is None:
        sheet_to_read = sheets[0]

    df = pd.read_excel(uploaded_file, sheet_name=sheet_to_read)

    # Filtra apenas os indicadores de interesse
    if RES_COL_INDICADOR_NOME not in df.columns:
        return pd.DataFrame()

    df = df[df[RES_COL_INDICADOR_NOME].isin(RES_INDICADORES_FILTRO)].copy()

    if df.empty:
        return df

    # Garante tipos numéricos
    for c in [RES_COL_VOLUME, RES_COL_INDICADOR_VAL, RES_COL_TMA, RES_COL_TMR]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")

    # Datas
    for c in [RES_COL_DT_INICIO, RES_COL_DT_FIM]:
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], errors="coerce")

    # ANOMES como string
    if RES_COL_ANOMES in df.columns:
        df[RES_COL_ANOMES] = df[RES_COL_ANOMES].astype(str).str.strip()

    # Coluna de aderência normalizada:
    # Para REPROGRAMAÇÃO GPON: INDICADOR=1 → NÃO ADERENTE, então invertemos
    df["ADERENTE"] = df.apply(
        lambda row: (
            (row[RES_COL_INDICADOR_VAL] == 0)
            if row[RES_COL_INDICADOR_NOME] in RES_IND_INVERTIDOS
            else (row[RES_COL_INDICADOR_VAL] == 1)
        ),
        axis=1,
    ).astype(int)

    # Campo de data para evolução diária
    if RES_COL_DT_INICIO in df.columns:
        df["DATA_DIA"] = df[RES_COL_DT_INICIO].dt.normalize()

    return df


def res_kpis_por_indicador(df: pd.DataFrame) -> pd.DataFrame:
    """
    Retorna uma linha por indicador com: Volume, Aderentes, Aderência %,
    TMA médio, TMR médio.
    """
    if df.empty:
        return pd.DataFrame()

    has_tma = RES_COL_TMA in df.columns
    has_tmr = RES_COL_TMR in df.columns

    agg = {
        RES_COL_VOLUME: "sum",
        "ADERENTE": "sum",
    }
    if has_tma:
        agg[RES_COL_TMA] = "mean"
    if has_tmr:
        agg[RES_COL_TMR] = "mean"

    g = df.groupby(RES_COL_INDICADOR_NOME).agg(agg).reset_index()
    g.columns = (
        ["Indicador", "Volume", "Aderentes"]
        + (["TMA_Medio"] if has_tma else [])
        + (["TMR_Medio"] if has_tmr else [])
    )
    g["Aderencia_Pct"] = (g["Aderentes"] / g["Volume"] * 100).round(1)

    # Ordena pela ordem configurada
    order = {ind: i for i, ind in enumerate(RES_INDICADORES_FILTRO)}
    g["_ord"] = g["Indicador"].map(order)
    g = g.sort_values("_ord").drop(columns="_ord").reset_index(drop=True)

    return g


def res_por_regional(df: pd.DataFrame, indicador: str | None = None) -> pd.DataFrame:
    """Breakdown por regional, opcionalmente filtrado por indicador."""
    if df.empty:
        return pd.DataFrame()
    sub = df[df[RES_COL_INDICADOR_NOME] == indicador] if indicador else df
    if sub.empty:
        return pd.DataFrame()
    g = sub.groupby(RES_COL_REGIONAL).agg(
        Volume=(RES_COL_VOLUME, "sum"),
        Aderentes=("ADERENTE", "sum"),
    ).reset_index().rename(columns={RES_COL_REGIONAL: "Regional"})
    g["Aderencia_Pct"] = (g["Aderentes"] / g["Volume"] * 100).round(1)
    return g.sort_values("Volume", ascending=False).reset_index(drop=True)


def res_por_natureza(df: pd.DataFrame, indicador: str | None = None) -> pd.DataFrame:
    """Breakdown por natureza, opcionalmente filtrado por indicador."""
    if df.empty or RES_COL_NATUREZA not in df.columns:
        return pd.DataFrame()
    sub = df[df[RES_COL_INDICADOR_NOME] == indicador] if indicador else df
    if sub.empty:
        return pd.DataFrame()
    g = sub.groupby(RES_COL_NATUREZA).agg(
        Volume=(RES_COL_VOLUME, "sum"),
        Aderentes=("ADERENTE", "sum"),
    ).reset_index().rename(columns={RES_COL_NATUREZA: "Natureza"})
    g["Aderencia_Pct"] = (g["Aderentes"] / g["Volume"] * 100).round(1)
    return g.sort_values("Volume", ascending=False).reset_index(drop=True)


def res_por_solucao(df: pd.DataFrame, indicador: str | None = None, top_n: int = 15) -> pd.DataFrame:
    """Top causas/soluções por volume."""
    if df.empty or RES_COL_SOLUCAO not in df.columns:
        return pd.DataFrame()
    sub = df[df[RES_COL_INDICADOR_NOME] == indicador] if indicador else df
    if sub.empty:
        return pd.DataFrame()
    g = sub.groupby(RES_COL_SOLUCAO).agg(
        Volume=(RES_COL_VOLUME, "sum"),
        Aderentes=("ADERENTE", "sum"),
    ).reset_index().rename(columns={RES_COL_SOLUCAO: "Solução"})
    g["Aderencia_Pct"] = (g["Aderentes"] / g["Volume"] * 100).round(1)
    return g.sort_values("Volume", ascending=False).head(top_n).reset_index(drop=True)


def res_por_impacto(df: pd.DataFrame, indicador: str | None = None) -> pd.DataFrame:
    """Breakdown por impacto (Massivo / Não Massivo)."""
    if df.empty or RES_COL_IMPACTO not in df.columns:
        return pd.DataFrame()
    sub = df[df[RES_COL_INDICADOR_NOME] == indicador] if indicador else df
    if sub.empty:
        return pd.DataFrame()
    g = sub.groupby(RES_COL_IMPACTO).agg(
        Volume=(RES_COL_VOLUME, "sum"),
        Aderentes=("ADERENTE", "sum"),
    ).reset_index().rename(columns={RES_COL_IMPACTO: "Impacto"})
    g["Aderencia_Pct"] = (g["Aderentes"] / g["Volume"] * 100).round(1)
    return g.sort_values("Volume", ascending=False).reset_index(drop=True)


def res_evolucao_diaria(df: pd.DataFrame, indicador: str | None = None) -> pd.DataFrame:
    """Evolução diária de volume e aderência."""
    if df.empty or "DATA_DIA" not in df.columns:
        return pd.DataFrame()
    sub = df[df[RES_COL_INDICADOR_NOME] == indicador] if indicador else df
    if sub.empty:
        return pd.DataFrame()
    sub = sub.dropna(subset=["DATA_DIA"])
    g = sub.groupby("DATA_DIA").agg(
        Volume=(RES_COL_VOLUME, "sum"),
        Aderentes=("ADERENTE", "sum"),
    ).reset_index().rename(columns={"DATA_DIA": "Data"})
    g["Aderencia_Pct"] = (g["Aderentes"] / g["Volume"] * 100).round(1)
    return g.sort_values("Data").reset_index(drop=True)


# =====================================================
# Funções originais (sem alteração)
# =====================================================
def resumo_mensal(df: pd.DataFrame) -> pd.DataFrame:
    """Gera resumo mensal por analista."""
    agg_dict = {
        COL_DATA: "count",
        COL_VOL_TOTAL: "sum",
    }

    for vc in VOL_COLS.keys():
        if vc in df.columns:
            agg_dict[vc] = "sum"

    group_cols = [COL_LOGIN, COL_NOME, "Setor", COL_MES, COL_ANOMES]
    existing_group = [c for c in group_cols if c in df.columns]

    g = df.groupby(existing_group).agg(agg_dict).reset_index()
    g = g.rename(columns={COL_DATA: "Dias_Trabalhados"})

    g["Media_Diaria"] = (g[COL_VOL_TOTAL] / g["Dias_Trabalhados"]).round(1)

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
