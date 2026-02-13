from __future__ import annotations

import streamlit as st
import pandas as pd

from src.config import EQUIPE, METAS
from src.parsers import (
    parse_chat_toa_from_toa,
    parse_etit_outage_sem_sinal_gpon_from_residencial,
    parse_log_reprog_gpon_from_residencial,
)

st.set_page_config(page_title="Dashboard Indicadores - Upload", layout="wide")

BASE = pd.DataFrame(EQUIPE, columns=["Matricula","Nome","Setor"])
EQUIPE_IDS = set(BASE["Matricula"].tolist())

def status_por_meta(valor_pct: float | None, meta: float, direcao: str) -> str:
    if valor_pct is None or (isinstance(valor_pct, float) and pd.isna(valor_pct)):
        return "—"
    if direcao == "up":
        return "✅ Dentro" if valor_pct >= meta else "❌ Fora"
    return "✅ Dentro" if valor_pct <= meta else "❌ Fora"

st.title("📊 Dashboard de Indicadores (Upload de Planilhas)")
st.caption("Upload das planilhas → Processar → Visão por analista (somente sua equipe).")

c1, c2 = st.columns(2)
with c1:
    f_toa = st.file_uploader("Planilha TOA (Analitico Indicadores TOA)", type=["xlsx","xls"], key="toa")
with c2:
    f_res = st.file_uploader("Planilha Residencial (Analítico Indicadores Residencial)", type=["xlsx","xls"], key="res")

if st.button("🚀 Processar Dados", use_container_width=True, type="primary"):
    resultados = []

    if f_toa:
        r = parse_chat_toa_from_toa(f_toa, EQUIPE_IDS)
        if r: resultados.append(r)

    if f_res:
        r1 = parse_etit_outage_sem_sinal_gpon_from_residencial(f_res, EQUIPE_IDS)
        if r1: resultados.append(r1)
        r2 = parse_log_reprog_gpon_from_residencial(f_res, EQUIPE_IDS)
        if r2: resultados.append(r2)

    # Monta tabela final wide
    df_final = BASE.copy()
    for r in resultados:
        col = r.indicador
        tmp = r.df[["Matricula","pct"]].rename(columns={"pct": col})
        df_final = df_final.merge(tmp, on="Matricula", how="left")

    # Status por meta
    for indicador, meta_info in METAS.items():
        if indicador in df_final.columns:
            df_final[indicador + " - Status"] = df_final[indicador].apply(
                lambda v: status_por_meta(v, meta_info["meta"], meta_info["direcao"])
            )

    st.subheader("✅ Visão por analista")
    st.dataframe(df_final, use_container_width=True)

    # Painel de alertas (fora da meta)
    st.subheader("🚨 Alertas (fora da meta)")
    alert_cols = []
    for indicador in METAS.keys():
        status_col = indicador + " - Status"
        if status_col in df_final.columns:
            alert_cols.append(status_col)

    if alert_cols:
        alert = df_final[df_final[alert_cols].apply(lambda row: any(v == "❌ Fora" for v in row), axis=1)]
        st.dataframe(alert[["Matricula","Nome","Setor"] + [c for c in df_final.columns if c.endswith("Status")]], use_container_width=True)
    else:
        st.info("Nenhum indicador processado ainda (verifique se os nomes existem nas planilhas).")

    # Ranking por indicador
    st.subheader("🏁 Ranking")
    indicador_sel = st.selectbox("Escolha um indicador para ranquear", [c for c in df_final.columns if c in METAS], index=0 if any(c in METAS for c in df_final.columns) else None)
    if indicador_sel:
        st.dataframe(
            df_final[["Matricula","Nome","Setor",indicador_sel, indicador_sel+" - Status"]].sort_values(indicador_sel, ascending=(METAS[indicador_sel]["direcao"]=="down")),
            use_container_width=True
        )
