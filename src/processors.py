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
    # DPA Ocupação
    DPA_MESES_PT, DPA_SHEET_ANALISTAS, DPA_SHEET_CONSOLIDADO,
    # Indicadores TOA
    TOA_IND_SHEET, TOA_INDICADORES_FILTRO, TOA_IND_INVERTIDOS,
    TOA_COL_INDICADOR_NOME, TOA_COL_LOGIN, TOA_COL_INDICADOR,
    TOA_COL_STATUS, TOA_COL_REGIONAL,
    TOA_COL_TIPO_ATIVIDADE, TOA_COL_REDE, TOA_COL_MERCADO,
    TOA_COL_NATUREZA, TOA_COL_SOLUCAO,
    TOA_COL_TMR, TOA_COL_AGING, TOA_COL_DATA,
    TOA_COL_DT_CANCELAMENTO, TOA_COL_DT_INICIO_FORM, TOA_COL_DT_FIM_FORM,
    TOA_COL_ANOMES, TOA_COL_ID_ATIVIDADE, TOA_AGING_ORDER,
    TOA_IND_CANCELADAS, TOA_IND_VALIDACAO,
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

    # Remove colunas sem cabeçalho (Unnamed ou NaN)
    df.columns = [str(c) for c in df.columns]
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
    return group.sort_values("Total_Eventos", ascending=False).reset_index(drop=True)


def etit_por_demanda(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame()
    return df.groupby(ETIT_COL_DEMANDA).agg(
        Eventos=(ETIT_COL_VOLUME, "sum"),
        Aderentes=(ETIT_COL_INDICADOR_VAL, "sum"),
        TMA_Medio=(ETIT_COL_TMA, "mean"),
        TMR_Medio=(ETIT_COL_TMR, "mean"),
    ).reset_index().rename(columns={ETIT_COL_DEMANDA: "Demanda"})


def etit_por_tipo(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame()
    return df.groupby(ETIT_COL_TIPO).agg(
        Eventos=(ETIT_COL_VOLUME, "sum"),
        Aderentes=(ETIT_COL_INDICADOR_VAL, "sum"),
    ).reset_index().rename(columns={ETIT_COL_TIPO: "Tipo"}).sort_values("Eventos", ascending=False)


def etit_por_causa(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame()
    return df.groupby(ETIT_COL_CAUSA).agg(
        Eventos=(ETIT_COL_VOLUME, "sum"),
        Aderentes=(ETIT_COL_INDICADOR_VAL, "sum"),
    ).reset_index().rename(columns={ETIT_COL_CAUSA: "Causa"}).sort_values("Eventos", ascending=False)


def etit_por_regional(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame()
    return df.groupby(ETIT_COL_REGIONAL).agg(
        Eventos=(ETIT_COL_VOLUME, "sum"),
        Aderentes=(ETIT_COL_INDICADOR_VAL, "sum"),
    ).reset_index().rename(columns={ETIT_COL_REGIONAL: "Regional"}).sort_values("Eventos", ascending=False)


def etit_por_turno(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame()
    return df.groupby(ETIT_COL_TURNO).agg(
        Eventos=(ETIT_COL_VOLUME, "sum"),
        Aderentes=(ETIT_COL_INDICADOR_VAL, "sum"),
    ).reset_index().rename(columns={ETIT_COL_TURNO: "Turno"}).sort_values("Eventos", ascending=False)


def etit_evolucao_diaria(df: pd.DataFrame) -> pd.DataFrame:
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

    if RES_COL_INDICADOR_NOME not in df.columns:
        return pd.DataFrame()

    df = df[df[RES_COL_INDICADOR_NOME].isin(RES_INDICADORES_FILTRO)].copy()

    if df.empty:
        return df

    for c in [RES_COL_VOLUME, RES_COL_INDICADOR_VAL, RES_COL_TMA, RES_COL_TMR]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")

    for c in [RES_COL_DT_INICIO, RES_COL_DT_FIM]:
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], errors="coerce")

    if RES_COL_ANOMES in df.columns:
        df[RES_COL_ANOMES] = df[RES_COL_ANOMES].astype(str).str.strip()

    df["ADERENTE"] = df.apply(
        lambda row: (
            (row[RES_COL_INDICADOR_VAL] == 0)
            if row[RES_COL_INDICADOR_NOME] in RES_IND_INVERTIDOS
            else (row[RES_COL_INDICADOR_VAL] == 1)
        ),
        axis=1,
    ).astype(int)

    if RES_COL_DT_INICIO in df.columns:
        df["DATA_DIA"] = df[RES_COL_DT_INICIO].dt.normalize()

    return df


def res_kpis_por_indicador(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame()
    has_tma = RES_COL_TMA in df.columns
    has_tmr = RES_COL_TMR in df.columns
    agg = {RES_COL_VOLUME: "sum", "ADERENTE": "sum"}
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
    order = {ind: i for i, ind in enumerate(RES_INDICADORES_FILTRO)}
    g["_ord"] = g["Indicador"].map(order)
    return g.sort_values("_ord").drop(columns="_ord").reset_index(drop=True)


def res_por_regional(df: pd.DataFrame, indicador=None) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame()
    sub = df[df[RES_COL_INDICADOR_NOME] == indicador] if indicador else df
    if sub.empty:
        return pd.DataFrame()
    g = sub.groupby(RES_COL_REGIONAL).agg(
        Volume=(RES_COL_VOLUME, "sum"), Aderentes=("ADERENTE", "sum"),
    ).reset_index().rename(columns={RES_COL_REGIONAL: "Regional"})
    g["Aderencia_Pct"] = (g["Aderentes"] / g["Volume"] * 100).round(1)
    return g.sort_values("Volume", ascending=False).reset_index(drop=True)


def res_por_natureza(df: pd.DataFrame, indicador=None) -> pd.DataFrame:
    if df.empty or RES_COL_NATUREZA not in df.columns:
        return pd.DataFrame()
    sub = df[df[RES_COL_INDICADOR_NOME] == indicador] if indicador else df
    if sub.empty:
        return pd.DataFrame()
    g = sub.groupby(RES_COL_NATUREZA).agg(
        Volume=(RES_COL_VOLUME, "sum"), Aderentes=("ADERENTE", "sum"),
    ).reset_index().rename(columns={RES_COL_NATUREZA: "Natureza"})
    g["Aderencia_Pct"] = (g["Aderentes"] / g["Volume"] * 100).round(1)
    return g.sort_values("Volume", ascending=False).reset_index(drop=True)


def res_por_solucao(df: pd.DataFrame, indicador=None, top_n=15) -> pd.DataFrame:
    if df.empty or RES_COL_SOLUCAO not in df.columns:
        return pd.DataFrame()
    sub = df[df[RES_COL_INDICADOR_NOME] == indicador] if indicador else df
    if sub.empty:
        return pd.DataFrame()
    g = sub.groupby(RES_COL_SOLUCAO).agg(
        Volume=(RES_COL_VOLUME, "sum"), Aderentes=("ADERENTE", "sum"),
    ).reset_index().rename(columns={RES_COL_SOLUCAO: "Solução"})
    g["Aderencia_Pct"] = (g["Aderentes"] / g["Volume"] * 100).round(1)
    return g.sort_values("Volume", ascending=False).head(top_n).reset_index(drop=True)


def res_por_impacto(df: pd.DataFrame, indicador=None) -> pd.DataFrame:
    if df.empty or RES_COL_IMPACTO not in df.columns:
        return pd.DataFrame()
    sub = df[df[RES_COL_INDICADOR_NOME] == indicador] if indicador else df
    if sub.empty:
        return pd.DataFrame()
    g = sub.groupby(RES_COL_IMPACTO).agg(
        Volume=(RES_COL_VOLUME, "sum"), Aderentes=("ADERENTE", "sum"),
    ).reset_index().rename(columns={RES_COL_IMPACTO: "Impacto"})
    g["Aderencia_Pct"] = (g["Aderentes"] / g["Volume"] * 100).round(1)
    return g.sort_values("Volume", ascending=False).reset_index(drop=True)


def res_evolucao_diaria(df: pd.DataFrame, indicador=None) -> pd.DataFrame:
    if df.empty or "DATA_DIA" not in df.columns:
        return pd.DataFrame()
    sub = df[df[RES_COL_INDICADOR_NOME] == indicador] if indicador else df
    if sub.empty:
        return pd.DataFrame()
    sub = sub.dropna(subset=["DATA_DIA"])
    g = sub.groupby("DATA_DIA").agg(
        Volume=(RES_COL_VOLUME, "sum"), Aderentes=("ADERENTE", "sum"),
    ).reset_index().rename(columns={"DATA_DIA": "Data"})
    g["Aderencia_Pct"] = (g["Aderentes"] / g["Volume"] * 100).round(1)
    return g.sort_values("Data").reset_index(drop=True)


# =====================================================
# OCUPAÇÃO DPA — Loader (planilha Ocupação_DPA_2026)
# =====================================================

def _dpa_detect_mes_recente(df_raw: pd.DataFrame) -> dict:
    """
    Varre a aba Consolidado (header=None) e retorna o último mês de 2026
    com DPA% > 0.
    Estrutura: col 26 = nome do mês, col 30 = % DPA 2026.
    """
    resultado = {"mes_nome": None, "mes_num": None, "dpa_geral_pct": None}
    for _, row in df_raw.iterrows():
        mes_val = str(row.get(26, "")).strip()
        if mes_val in DPA_MESES_PT:
            try:
                pct_f = float(row.get(30, None))
                if pct_f > 0:
                    resultado = {
                        "mes_nome": mes_val,
                        "mes_num": DPA_MESES_PT.index(mes_val) + 1,
                        "dpa_geral_pct": round(pct_f * 100, 2),
                    }
            except (TypeError, ValueError):
                pass
    return resultado


def _dpa_extract_analistas(df_raw: pd.DataFrame) -> pd.DataFrame:
    """
    Extrai tabela de DPA por analista da aba Analistas (header=None).
    Pivot está na col 26 (Login), col 27 (Ocupação Produtiva), col 28 (% Produtivo).
    Procura a linha de header dinamicamente.
    """
    header_row = None
    for i, row in df_raw.iterrows():
        if str(row.get(26, "")).strip() == "Rótulos de Linha":
            header_row = i
            break
    if header_row is None:
        return pd.DataFrame()

    rows = []
    skip_tokens = {"nan", "Total Geral", "COP REDE RJ", "", "Rótulos de Linha"}
    for i in range(header_row + 1, len(df_raw)):
        row = df_raw.iloc[i]
        login = str(row.get(26, "")).strip()
        if not login or login in skip_tokens:
            continue
        try:
            pct_f = float(row.get(28, None))
            rows.append({"Login": login, "DPA_Pct_Oficial": round(pct_f * 100, 2)})
        except (TypeError, ValueError):
            pass

    if not rows:
        return pd.DataFrame()

    df = pd.DataFrame(rows)
    # Merge com nome e setor via BASE_EQUIPE
    df = df.merge(
        BASE_EQUIPE[["Matricula", "Nome", "Setor"]],
        left_on="Login", right_on="Matricula", how="inner",  # inner = somente equipe
    ).drop(columns="Matricula")
    df = df.sort_values("DPA_Pct_Oficial", ascending=False).reset_index(drop=True)
    df.index += 1
    df.index.name = "#"
    return df


def load_dpa_ocupacao(uploaded_file) -> tuple[pd.DataFrame, dict]:
    """
    Carrega a planilha Ocupação DPA 2026 e retorna:
    - df_analistas: DataFrame com Login, Nome, Setor, DPA_Pct_Oficial
    - mes_info: dict com mes_nome, mes_num, dpa_geral_pct
    """
    if hasattr(uploaded_file, "seek"):
        uploaded_file.seek(0)

    df_consolidado = pd.read_excel(uploaded_file, sheet_name=DPA_SHEET_CONSOLIDADO, header=None)

    if hasattr(uploaded_file, "seek"):
        uploaded_file.seek(0)

    df_analistas_raw = pd.read_excel(uploaded_file, sheet_name=DPA_SHEET_ANALISTAS, header=None)

    mes_info = _dpa_detect_mes_recente(df_consolidado)
    df_analistas = _dpa_extract_analistas(df_analistas_raw)

    return df_analistas, mes_info


def dpa_ranking(df_analistas: pd.DataFrame) -> pd.DataFrame:
    """Retorna DataFrame de ranking de DPA já ordenado."""
    if df_analistas.empty:
        return pd.DataFrame()
    df = df_analistas.copy()
    df["Nome_Curto"] = df["Nome"].apply(primeiro_nome)
    return df[["Nome_Curto", "Login", "Setor", "DPA_Pct_Oficial"]].rename(
        columns={"Nome_Curto": "Analista", "DPA_Pct_Oficial": "DPA %"}
    )


def dpa_comparativo(df_analistas: pd.DataFrame, resumo_prod: pd.DataFrame) -> pd.DataFrame:
    """
    Junta DPA Oficial (planilha Ocupação) com DPA calculado (planilha Produtividade).
    Retorna DataFrame com ambos para comparação.
    """
    if df_analistas.empty or resumo_prod.empty:
        return pd.DataFrame()

    oficial = df_analistas[["Login", "Nome", "DPA_Pct_Oficial"]].copy()

    # Verifica se DPA_Media existe no resumo
    if "DPA_Media" not in resumo_prod.columns:
        return oficial

    calculado = resumo_prod[["USUARIO_LOGIN", "DPA_Media"]].rename(
        columns={"USUARIO_LOGIN": "Login", "DPA_Media": "DPA_Calculado"}
    )

    merged = oficial.merge(calculado, on="Login", how="left")
    merged["Diferença"] = (merged["DPA_Pct_Oficial"] - merged["DPA_Calculado"]).round(2)
    merged["Nome_Curto"] = merged["Nome"].apply(primeiro_nome)
    return merged[["Nome_Curto", "Login", "DPA_Pct_Oficial", "DPA_Calculado", "Diferença"]].rename(
        columns={"Nome_Curto": "Analista", "DPA_Pct_Oficial": "DPA Oficial %", "DPA_Calculado": "DPA Calculado %"}
    ).sort_values("DPA Oficial %", ascending=False).reset_index(drop=True)


# =====================================================
# Funções originais de produtividade
# =====================================================
def resumo_mensal(df: pd.DataFrame) -> pd.DataFrame:
    agg_dict = {COL_DATA: "count", COL_VOL_TOTAL: "sum"}
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
    agg_dict = {COL_DATA: "count", COL_VOL_TOTAL: "sum"}
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
    if COL_DATA not in df.columns:
        return pd.DataFrame()
    daily = df.groupby(COL_DATA).agg(
        Vol_Total=(COL_VOL_TOTAL, "sum"),
        Analistas=(COL_LOGIN, "nunique"),
    ).reset_index()
    daily["Media_por_Analista"] = (daily["Vol_Total"] / daily["Analistas"]).round(1)
    return daily


def composicao_volume(df: pd.DataFrame) -> pd.DataFrame:
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
    parts = str(nome_completo).strip().split()
    if len(parts) <= 2:
        return nome_completo
    return f"{parts[0]} {parts[-1]}"


# =====================================================
# INDICADORES TOA — Loader e processadores
# =====================================================

def load_toa_indicadores(uploaded_file) -> pd.DataFrame:
    """
    Lê a planilha Analitico_Indicadores_TOA e retorna apenas
    TAREFAS CANCELADAS e TEMPO DE VALIDAÇÃO DO FORMULÁRIO
    filtrados pela equipe monitorada.

    Detecção automática do mês mais recente via coluna ANOMES.
    """
    if hasattr(uploaded_file, "seek"):
        uploaded_file.seek(0)

    df = pd.read_excel(uploaded_file, sheet_name=TOA_IND_SHEET)

    # Filtrar indicadores de interesse
    if TOA_COL_INDICADOR_NOME not in df.columns:
        return pd.DataFrame()

    df = df[df[TOA_COL_INDICADOR_NOME].isin(TOA_INDICADORES_FILTRO)].copy()
    if df.empty:
        return df

    # Normalizar login (maiúsculo, sem espaços)
    df[TOA_COL_LOGIN] = df[TOA_COL_LOGIN].astype(str).str.strip().str.upper()

    # Detectar ANOMES mais recente e filtrar
    if TOA_COL_ANOMES in df.columns:
        df[TOA_COL_ANOMES] = pd.to_numeric(df[TOA_COL_ANOMES], errors="coerce")
        anomes_recente = df[TOA_COL_ANOMES].max()
        df = df[df[TOA_COL_ANOMES] == anomes_recente].copy()

    # Filtrar equipe
    df = df[df[TOA_COL_LOGIN].isin(EQUIPE_IDS)].copy()
    if df.empty:
        return df

    # Merge com nome e setor
    df = df.merge(
        BASE_EQUIPE[["Matricula", "Nome", "Setor"]],
        left_on=TOA_COL_LOGIN, right_on="Matricula", how="left"
    )

    # Tipos
    if TOA_COL_INDICADOR in df.columns:
        df[TOA_COL_INDICADOR] = pd.to_numeric(df[TOA_COL_INDICADOR], errors="coerce")

    for c in [TOA_COL_DATA, TOA_COL_DT_CANCELAMENTO, TOA_COL_DT_INICIO_FORM, TOA_COL_DT_FIM_FORM]:
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], errors="coerce")

    # TMR já vem como timedelta; converter para minutos para facilitar análise
    if TOA_COL_TMR in df.columns:
        try:
            if pd.api.types.is_timedelta64_dtype(df[TOA_COL_TMR]):
                df["TMR_min"] = df[TOA_COL_TMR].dt.total_seconds() / 60
            else:
                df[TOA_COL_TMR] = pd.to_timedelta(df[TOA_COL_TMR], errors="coerce")
                df["TMR_min"] = df[TOA_COL_TMR].dt.total_seconds() / 60
        except Exception:
            df["TMR_min"] = pd.to_numeric(df[TOA_COL_TMR], errors="coerce")

    # Coluna ADERENTE normalizada:
    # Canceladas: INDICADOR=1 → NÃO ADERENTE → invertemos
    # Validação:  INDICADOR=1 → ADERENTE
    df["ADERENTE"] = df.apply(
        lambda row: (
            (row[TOA_COL_INDICADOR] == 0)
            if row[TOA_COL_INDICADOR_NOME] in TOA_IND_INVERTIDOS
            else (row[TOA_COL_INDICADOR] == 1)
        ),
        axis=1,
    ).astype(int)

    # Data para evolução diária
    if TOA_COL_DATA in df.columns:
        df["DATA_DIA"] = df[TOA_COL_DATA].dt.normalize()

    return df


def toa_anomes_recente(df: pd.DataFrame) -> int | None:
    """Retorna o ANOMES mais recente presente no DataFrame."""
    if df.empty or TOA_COL_ANOMES not in df.columns:
        return None
    v = pd.to_numeric(df[TOA_COL_ANOMES], errors="coerce").max()
    return int(v) if pd.notna(v) else None


def toa_resumo_por_indicador(df: pd.DataFrame) -> pd.DataFrame:
    """KPI geral por indicador: total, aderentes, aderência%, TMR médio."""
    if df.empty:
        return pd.DataFrame()
    rows = []
    for ind in TOA_INDICADORES_FILTRO:
        sub = df[df[TOA_COL_INDICADOR_NOME] == ind]
        if sub.empty:
            continue
        total = len(sub)
        ader  = int(sub["ADERENTE"].sum())
        pct   = round(ader / total * 100, 1) if total > 0 else 0.0
        tmr_m = sub["TMR_min"].mean() if "TMR_min" in sub.columns else None
        rows.append({
            "Indicador": ind,
            "Total": total,
            "Aderentes": ader,
            "Aderencia_Pct": pct,
            "TMR_Medio_min": round(tmr_m, 2) if tmr_m is not None and pd.notna(tmr_m) else None,
        })
    return pd.DataFrame(rows)


# ---- TAREFAS CANCELADAS ----

def toa_canceladas_por_analista(df: pd.DataFrame) -> pd.DataFrame:
    """
    Ranking de tarefas canceladas por analista.
    Canceladas = todas as linhas deste indicador (INDICADOR=1 sempre).
    Menor = melhor.
    """
    sub = df[df[TOA_COL_INDICADOR_NOME] == TOA_IND_CANCELADAS].copy()
    if sub.empty:
        return pd.DataFrame()
    g = sub.groupby([TOA_COL_LOGIN, "Nome", "Setor"]).agg(
        Canceladas=(TOA_COL_INDICADOR, "count"),
        TMR_Medio_h=("TMR_min", lambda x: round(x.mean() / 60, 2) if x.notna().any() else None),
    ).reset_index().rename(columns={TOA_COL_LOGIN: "Login"})
    return g.sort_values("Canceladas", ascending=False).reset_index(drop=True)


def toa_canceladas_por_tipo(df: pd.DataFrame) -> pd.DataFrame:
    """Breakdown de tarefas canceladas por TIPO_ATIVIDADE."""
    sub = df[df[TOA_COL_INDICADOR_NOME] == TOA_IND_CANCELADAS]
    if sub.empty or TOA_COL_TIPO_ATIVIDADE not in sub.columns:
        return pd.DataFrame()
    g = sub.groupby(TOA_COL_TIPO_ATIVIDADE).size().reset_index(name="Canceladas")
    g.columns = ["Tipo Atividade", "Canceladas"]
    return g.sort_values("Canceladas", ascending=False).reset_index(drop=True)


def toa_canceladas_por_aging(df: pd.DataFrame) -> pd.DataFrame:
    """Distribuição de tarefas canceladas por faixa de AGING."""
    sub = df[df[TOA_COL_INDICADOR_NOME] == TOA_IND_CANCELADAS]
    if sub.empty or TOA_COL_AGING not in sub.columns:
        return pd.DataFrame()
    g = sub.groupby(TOA_COL_AGING).size().reset_index(name="Canceladas")
    g.columns = ["Aging", "Canceladas"]
    # Ordenar pela ordem definida em config
    order_map = {v: i for i, v in enumerate(TOA_AGING_ORDER)}
    g["_ord"] = g["Aging"].map(order_map).fillna(99)
    return g.sort_values("_ord").drop(columns="_ord").reset_index(drop=True)


def toa_canceladas_por_rede(df: pd.DataFrame) -> pd.DataFrame:
    """Breakdown de tarefas canceladas por REDE."""
    sub = df[df[TOA_COL_INDICADOR_NOME] == TOA_IND_CANCELADAS]
    if sub.empty or TOA_COL_REDE not in sub.columns:
        return pd.DataFrame()
    g = sub.groupby(TOA_COL_REDE).size().reset_index(name="Canceladas")
    g.columns = ["Rede", "Canceladas"]
    return g.sort_values("Canceladas", ascending=False).reset_index(drop=True)


def toa_canceladas_por_regional(df: pd.DataFrame) -> pd.DataFrame:
    """Breakdown de tarefas canceladas por Regional."""
    sub = df[df[TOA_COL_INDICADOR_NOME] == TOA_IND_CANCELADAS]
    if sub.empty or TOA_COL_REGIONAL not in sub.columns:
        return pd.DataFrame()
    g = sub.groupby(TOA_COL_REGIONAL).size().reset_index(name="Canceladas")
    g.columns = ["Regional", "Canceladas"]
    return g.sort_values("Canceladas", ascending=False).reset_index(drop=True)


def toa_canceladas_evolucao(df: pd.DataFrame) -> pd.DataFrame:
    """Evolução diária de tarefas canceladas."""
    sub = df[df[TOA_COL_INDICADOR_NOME] == TOA_IND_CANCELADAS]
    if sub.empty or "DATA_DIA" not in sub.columns:
        return pd.DataFrame()
    g = sub.groupby("DATA_DIA").size().reset_index(name="Canceladas")
    g.columns = ["Data", "Canceladas"]
    return g.sort_values("Data").reset_index(drop=True)


# ---- TEMPO DE VALIDAÇÃO DO FORMULÁRIO ----

def toa_validacao_por_analista(df: pd.DataFrame) -> pd.DataFrame:
    """
    Ranking de aderência ao tempo de validação do formulário por analista.
    Inclui total de formulários, aderentes, aderência% e TMR médio em minutos.
    """
    sub = df[df[TOA_COL_INDICADOR_NOME] == TOA_IND_VALIDACAO].copy()
    if sub.empty:
        return pd.DataFrame()
    g = sub.groupby([TOA_COL_LOGIN, "Nome", "Setor"]).agg(
        Total=(TOA_COL_INDICADOR, "count"),
        Aderentes=("ADERENTE", "sum"),
        TMR_Medio_min=("TMR_min", "mean"),
    ).reset_index().rename(columns={TOA_COL_LOGIN: "Login"})
    g["Aderencia_Pct"] = (g["Aderentes"] / g["Total"] * 100).round(1)
    g["TMR_Medio_min"]  = g["TMR_Medio_min"].round(2)
    return g.sort_values("Aderencia_Pct", ascending=False).reset_index(drop=True)


def toa_validacao_por_tipo(df: pd.DataFrame) -> pd.DataFrame:
    """Breakdown do tempo de validação por TIPO_ATIVIDADE."""
    sub = df[df[TOA_COL_INDICADOR_NOME] == TOA_IND_VALIDACAO]
    if sub.empty or TOA_COL_TIPO_ATIVIDADE not in sub.columns:
        return pd.DataFrame()
    g = sub.groupby(TOA_COL_TIPO_ATIVIDADE).agg(
        Total=(TOA_COL_INDICADOR, "count"),
        Aderentes=("ADERENTE", "sum"),
        TMR_Medio_min=("TMR_min", "mean"),
    ).reset_index().rename(columns={TOA_COL_TIPO_ATIVIDADE: "Tipo Atividade"})
    g["Aderencia_Pct"] = (g["Aderentes"] / g["Total"] * 100).round(1)
    g["TMR_Medio_min"]  = g["TMR_Medio_min"].round(2)
    return g.sort_values("Total", ascending=False).reset_index(drop=True)


def toa_validacao_por_rede(df: pd.DataFrame) -> pd.DataFrame:
    """Breakdown do tempo de validação por REDE."""
    sub = df[df[TOA_COL_INDICADOR_NOME] == TOA_IND_VALIDACAO]
    if sub.empty or TOA_COL_REDE not in sub.columns:
        return pd.DataFrame()
    g = sub.groupby(TOA_COL_REDE).agg(
        Total=(TOA_COL_INDICADOR, "count"),
        Aderentes=("ADERENTE", "sum"),
        TMR_Medio_min=("TMR_min", "mean"),
    ).reset_index().rename(columns={TOA_COL_REDE: "Rede"})
    g["Aderencia_Pct"] = (g["Aderentes"] / g["Total"] * 100).round(1)
    g["TMR_Medio_min"]  = g["TMR_Medio_min"].round(2)
    return g.sort_values("Total", ascending=False).reset_index(drop=True)


def toa_validacao_por_regional(df: pd.DataFrame) -> pd.DataFrame:
    """Breakdown do tempo de validação por Regional."""
    sub = df[df[TOA_COL_INDICADOR_NOME] == TOA_IND_VALIDACAO]
    if sub.empty or TOA_COL_REGIONAL not in sub.columns:
        return pd.DataFrame()
    g = sub.groupby(TOA_COL_REGIONAL).agg(
        Total=(TOA_COL_INDICADOR, "count"),
        Aderentes=("ADERENTE", "sum"),
        TMR_Medio_min=("TMR_min", "mean"),
    ).reset_index().rename(columns={TOA_COL_REGIONAL: "Regional"})
    g["Aderencia_Pct"] = (g["Aderentes"] / g["Total"] * 100).round(1)
    g["TMR_Medio_min"]  = g["TMR_Medio_min"].round(2)
    return g.sort_values("Total", ascending=False).reset_index(drop=True)


def toa_validacao_evolucao(df: pd.DataFrame) -> pd.DataFrame:
    """Evolução diária da aderência ao tempo de validação."""
    sub = df[df[TOA_COL_INDICADOR_NOME] == TOA_IND_VALIDACAO]
    if sub.empty or "DATA_DIA" not in sub.columns:
        return pd.DataFrame()
    sub = sub.dropna(subset=["DATA_DIA"])
    g = sub.groupby("DATA_DIA").agg(
        Total=(TOA_COL_INDICADOR, "count"),
        Aderentes=("ADERENTE", "sum"),
        TMR_Medio_min=("TMR_min", "mean"),
    ).reset_index().rename(columns={"DATA_DIA": "Data"})
    g["Aderencia_Pct"] = (g["Aderentes"] / g["Total"] * 100).round(1)
    g["TMR_Medio_min"]  = g["TMR_Medio_min"].round(2)
    return g.sort_values("Data").reset_index(drop=True)
