import io
import traceback
import pandas as pd
import streamlit as st
import openpyxl

from src.config import BASE_EQUIPE, EQUIPE_IDS, METAS
from src.parsers import (
    parse_chat_toa_from_df,
    parse_gpon_from_residencial_df,
    status_por_meta,
)

st.set_page_config(page_title="Dashboard Indicadores (Upload)", layout="wide")


# -------------------------
# Helpers (debug)
# -------------------------
def list_sheets(uploaded_file):
    data = uploaded_file.getvalue()
    wb = openpyxl.load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    return wb.sheetnames


def safe_read_excel(uploaded_file, preferred_sheets=None):
    """
    Lê Excel com fallback:
    - mostra abas disponíveis
    - tenta pelas abas preferidas
    - se não achar, lê a primeira aba
    """
    sheets = list_sheets(uploaded_file)
    st.write("✅ Abas encontradas:", sheets)

    if preferred_sheets:
        for sh in preferred_sheets:
            if sh in sheets:
                st.write(f"➡️ Lendo aba: {sh}")
                return pd.read_excel(uploaded_file, sheet_name=sh)

    st.write(f"➡️ Lendo primeira aba: {sheets[0]}")
    return pd.read_excel(uploaded_file, sheet_name=sheets[0])


def show_df_info(df, title="Preview"):
    st.write(f"📌 {title} — linhas: {len(df)} | colunas: {len(df.columns)}")
    st.write("🔎 Colunas:", list(df.columns))
    st.dataframe(df.head(20), use_container_width=True)


# -------------------------
# UI
# -------------------------
st.title("📊 Dashboard de Indicadores (Upload de Planilhas)")
st.caption("Upload das planilhas → Processar → Visão por analista (somente sua equipe).")

colA, colB = st.columns(2)
with colA:
    f_toa = st.file_uploader(
        "Planilha TOA (Analitico Indicadores TOA)",
        type=["xlsx", "xls"],
        key="toa",
    )
with colB:
    f_res = st.file_uploader(
        "Planilha Residencial (Analítico Indicadores Residencial)",
        type=["xlsx", "xls"],
        key="res",
    )

st.markdown("---")

if st.button("🚀 Processar Dados", use_container_width=True):
    if (f_toa is None) or (f_res is None):
        st.error("Envie as duas planilhas (TOA e Residencial) antes de processar.")
        st.stop()

    try:
        with st.spinner("Lendo planilhas e calculando indicadores..."):
            # --------- TOA ----------
            df_toa = safe_read_excel(f_toa, preferred_sheets=["TOA", "Toa", "toa"])
            show_df_info(df_toa, "TOA (bruto)")

            # --------- RESIDENCIAL ----------
            df_res = safe_read_excel(
                f_res, preferred_sheets=["Analitico", "Analítico", "ANALITICO", "ANALÍTICO"]
            )
            show_df_info(df_res, "RESIDENCIAL (bruto)")

            # --------- PARSERS ----------
            chat_df = parse_chat_toa_from_df(df_toa, equipe_ids=EQUIPE_IDS)
            gpon_df = parse_gpon_from_residencial_df(df_res, equipe_ids=EQUIPE_IDS)

            # --------- MERGE FINAL ----------
            final = BASE_EQUIPE.copy()

            # chat_df: Matricula, Chat_TOA_pct
            final = final.merge(chat_df, on="Matricula", how="left")

            # gpon_df: Matricula, ETIT_GPON_pct, REPROG_GPON_pct
            final = final.merge(gpon_df, on="Matricula", how="left")

            # --------- STATUS vs META ----------
            # Chat TOA (>= 75)
            final["Chat_TOA_status"] = final["Chat_TOA_pct"].apply(
                lambda v: status_por_meta(v, METAS["Chat TOA"]["meta"], METAS["Chat TOA"]["direcao"])
            )

            # ETIT GPON (>= 90)
            final["ETIT_GPON_status"] = final["ETIT_GPON_pct"].apply(
                lambda v: status_por_meta(
                    v, METAS["ETIT Outage Sem Sinal (GPON)"]["meta"], METAS["ETIT Outage Sem Sinal (GPON)"]["direcao"]
                )
            )

            # Reprog GPON (<= 10)  -> aqui é % de NÃO aderente (reprog)
            final["REPROG_GPON_status"] = final["REPROG_GPON_pct"].apply(
                lambda v: status_por_meta(
                    v, METAS["Log Outage Reprog. GPON"]["meta"], METAS["Log Outage Reprog. GPON"]["direcao"]
                )
            )

        st.success("✅ Processamento concluído!")

        st.subheader("📌 Visão por analista (somente sua equipe)")
        st.dataframe(final, use_container_width=True)

        st.subheader("🏆 Ranking (Chat TOA)")
        rank_chat = final[["Matricula", "Nome", "Setor", "Chat_TOA_pct", "Chat_TOA_status"]].sort_values(
            "Chat_TOA_pct", ascending=False, na_position="last"
        )
        st.dataframe(rank_chat, use_container_width=True)

        st.subheader("🏆 Ranking (ETIT GPON)")
        rank_gpon = final[["Matricula", "Nome", "Setor", "ETIT_GPON_pct", "ETIT_GPON_status"]].sort_values(
            "ETIT_GPON_pct", ascending=False, na_position="last"
        )
        st.dataframe(rank_gpon, use_container_width=True)

        st.subheader("🚨 Ranking (Reprog GPON) — menor é melhor")
        rank_rep = final[["Matricula", "Nome", "Setor", "REPROG_GPON_pct", "REPROG_GPON_status"]].sort_values(
            "REPROG_GPON_pct", ascending=True, na_position="last"
        )
        st.dataframe(rank_rep, use_container_width=True)

        st.markdown("---")
        st.caption("Se algum indicador ficar em branco, o debug acima mostra as colunas/abas para ajustarmos o parser.")

    except Exception as e:
        st.error("❌ Erro ao processar as planilhas.")
        st.code(str(e))
        st.code(traceback.format_exc())
