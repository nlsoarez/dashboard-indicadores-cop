import traceback
import io
import pandas as pd
import numpy as np
import streamlit as st

from src.config import (
    BASE_EQUIPE, EQUIPE_IDS, VOL_COLS,
    VOL_COLS_RESIDENCIAL, VOL_COLS_EMPRESARIAL, VOL_COLS_AMBOS,
    COL_LOGIN, COL_NOME, COL_BASE, COL_DATA, COL_MES, COL_ANOMES,
    COL_VOL_TOTAL, COL_DPA_RESULTADO,
    COR_PRIMARIA, COR_SUCESSO, COR_ALERTA, COR_PERIGO, COR_INFO,
)
from src.processors import (
    load_produtividade, resumo_mensal, resumo_geral,
    evolucao_diaria, composicao_volume, primeiro_nome,
)

# =====================================================
# PAGE CONFIG
# =====================================================
st.set_page_config(
    page_title="Produtividade COP Rede",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# =====================================================
# CSS CUSTOMIZADO
# =====================================================
st.markdown("""
<style>
    /* Header */
    .main-header {
        background: linear-gradient(135deg, #1B4F72 0%, #2980B9 100%);
        padding: 1.5rem 2rem;
        border-radius: 12px;
        margin-bottom: 1.5rem;
        color: white;
    }
    .main-header h1 { color: white; margin: 0; font-size: 1.8rem; }
    .main-header p { color: #D6EAF8; margin: 0.3rem 0 0 0; font-size: 0.95rem; }

    /* KPI Cards */
    .kpi-card {
        background: white;
        border-radius: 10px;
        padding: 1.2rem;
        box-shadow: 0 2px 8px rgba(0,0,0,0.08);
        border-left: 4px solid;
        text-align: center;
    }
    .kpi-card .kpi-value {
        font-size: 2rem;
        font-weight: 700;
        margin: 0.3rem 0;
    }
    .kpi-card .kpi-label {
        font-size: 0.82rem;
        color: #7F8C8D;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    .kpi-card .kpi-delta {
        font-size: 0.8rem;
        margin-top: 0.2rem;
    }

    /* Section headers */
    .section-header {
        font-size: 1.2rem;
        font-weight: 600;
        color: #1B4F72;
        border-bottom: 2px solid #2980B9;
        padding-bottom: 0.4rem;
        margin: 1.5rem 0 1rem 0;
    }

    /* Ranking badges */
    .rank-badge {
        display: inline-block;
        width: 28px;
        height: 28px;
        border-radius: 50%;
        text-align: center;
        line-height: 28px;
        font-weight: 700;
        font-size: 0.85rem;
        color: white;
    }
    .rank-1 { background: #F1C40F; color: #333; }
    .rank-2 { background: #BDC3C7; color: #333; }
    .rank-3 { background: #CD6155; }

    /* Table styling */
    .dataframe { font-size: 0.85rem !important; }

    /* Performance highlight cards */
    .perf-card {
        padding: 0.9rem 1rem;
        border-radius: 10px;
        border-left: 4px solid;
        margin-bottom: 0.5rem;
    }
    .perf-best {
        background: linear-gradient(135deg, #d5f5e3 0%, #abebc6 100%);
        border-left-color: #27AE60;
    }
    .perf-worst {
        background: linear-gradient(135deg, #fadbd8 0%, #f5b7b1 100%);
        border-left-color: #E74C3C;
    }
    .perf-dpa {
        background: linear-gradient(135deg, #d6eaf8 0%, #aed6f1 100%);
        border-left-color: #2980B9;
    }

    /* Insight cards */
    .insight-card {
        background: white;
        border-radius: 10px;
        padding: 0.9rem 1rem;
        box-shadow: 0 2px 8px rgba(0,0,0,0.06);
        margin-bottom: 0.6rem;
        border-left: 4px solid #2980B9;
    }
    .tag-green {
        background: #d5f5e3; color: #1e8449;
        padding: 1px 7px; border-radius: 10px;
        font-size: 0.75rem; display: inline-block; margin: 1px 2px;
    }
    .tag-red {
        background: #fadbd8; color: #922b21;
        padding: 1px 7px; border-radius: 10px;
        font-size: 0.75rem; display: inline-block; margin: 1px 2px;
    }
    .sector-badge {
        background: #eaf2f8; color: #2c3e50;
        padding: 1px 7px; border-radius: 8px;
        font-size: 0.72rem; margin-left: 6px;
    }
    .rank-pill {
        background: #f0f0f0; color: #555;
        padding: 1px 7px; border-radius: 8px;
        font-size: 0.72rem; margin-left: 3px;
    }

    /* Sidebar */
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #1B4F72 0%, #154360 100%);
    }
    [data-testid="stSidebar"] * { color: white !important; }
    [data-testid="stSidebar"] .stSelectbox label { color: #D6EAF8 !important; }
</style>
""", unsafe_allow_html=True)


# =====================================================
# HELPERS
# =====================================================
def kpi_card(label, value, color, delta=None, suffix=""):
    delta_html = ""
    if delta is not None:
        delta_color = COR_SUCESSO if delta >= 0 else COR_PERIGO
        delta_icon = "▲" if delta >= 0 else "▼"
        delta_html = f'<div class="kpi-delta" style="color:{delta_color}">{delta_icon} {abs(delta):.1f}{suffix}</div>'
    return f"""
    <div class="kpi-card" style="border-left-color: {color};">
        <div class="kpi-label">{label}</div>
        <div class="kpi-value" style="color: {color};">{value}{suffix}</div>
        {delta_html}
    </div>
    """


def format_ranking(df_rank, value_col, label, ascending=False, suffix="%"):
    """Formata ranking com badges para top 3."""
    df_sorted = df_rank.sort_values(value_col, ascending=ascending, na_position="last").reset_index(drop=True)
    df_sorted.index = df_sorted.index + 1
    df_sorted.index.name = "#"
    return df_sorted


def get_sector_vol_cols(setor, available_cols):
    """Returns {col_name: display_label} for the given sector filter."""
    cols = {}
    if setor in ("Todos", "RESIDENCIAL"):
        cols.update(VOL_COLS_RESIDENCIAL)
    if setor in ("Todos", "EMPRESARIAL"):
        cols.update(VOL_COLS_EMPRESARIAL)
    cols.update(VOL_COLS_AMBOS)
    return {k: v for k, v in cols.items() if k in available_cols}


# =====================================================
# HEADER
# =====================================================
st.markdown("""
<div class="main-header">
    <h1>📊 Dashboard de Produtividade — COP Rede</h1>
    <p>Análise de produtividade da equipe · Upload da planilha analítica</p>
</div>
""", unsafe_allow_html=True)


# =====================================================
# UPLOAD (persiste dados no session_state)
# =====================================================
with st.container():
    col_upload, col_info = st.columns([2, 1])
    with col_upload:
        uploaded_file = st.file_uploader(
            "📁 Faça upload da planilha de Produtividade COP Rede (Analítico)",
            type=["xlsx", "xls"],
            help="Planilha com aba 'Analítico Produtividade 2026' ou similar",
        )
    with col_info:
        st.info(
            f"**Equipe monitorada:** {len(EQUIPE_IDS)} analistas\n\n"
            f"Empresarial: {len(BASE_EQUIPE[BASE_EQUIPE['Setor']=='EMPRESARIAL'])} · "
            f"Residencial: {len(BASE_EQUIPE[BASE_EQUIPE['Setor']=='RESIDENCIAL'])}"
        )

# Cache uploaded file in session_state to survive reruns
if uploaded_file is not None:
    st.session_state["uploaded_bytes"] = uploaded_file.getvalue()
    st.session_state["uploaded_name"] = uploaded_file.name

if "uploaded_bytes" not in st.session_state:
    st.markdown("---")
    st.markdown("### 👋 Bem-vindo!")
    st.markdown(
        "Faça upload da planilha **Produtividade COP Rede - Analítico** acima para "
        "visualizar os dados de produtividade da sua equipe."
    )

    with st.expander("📋 Analistas monitorados"):
        st.dataframe(BASE_EQUIPE, use_container_width=True, hide_index=True)
    st.stop()


# =====================================================
# PROCESSAR DADOS
# =====================================================
try:
    with st.spinner("Carregando e processando dados..."):
        file_obj = io.BytesIO(st.session_state["uploaded_bytes"])
        df = load_produtividade(file_obj)

    if df.empty:
        st.error("Nenhum analista da equipe encontrado na planilha. Verifique o arquivo.")
        st.stop()

except Exception as e:
    st.error(f"Erro ao processar a planilha: {e}")
    with st.expander("Detalhes do erro"):
        st.code(traceback.format_exc())
    st.stop()


# =====================================================
# SIDEBAR - FILTROS
# =====================================================
with st.sidebar:
    st.markdown("### 🔧 Filtros")

    # Meses disponíveis
    meses_disponiveis = sorted(df[COL_ANOMES].dropna().unique().tolist())
    meses_labels = df.drop_duplicates(COL_ANOMES).set_index(COL_ANOMES)[COL_MES].to_dict()

    mes_selecionado = st.selectbox(
        "Período",
        options=["Todos"] + meses_disponiveis,
        format_func=lambda x: "Todos os meses" if x == "Todos" else f"{meses_labels.get(x, x)} ({x})",
    )

    setor_selecionado = st.selectbox(
        "Setor",
        options=["Todos", "EMPRESARIAL", "RESIDENCIAL"],
    )

    st.markdown("---")
    if st.button("🗑️ Limpar dados carregados", use_container_width=True):
        for key in ["uploaded_bytes", "uploaded_name"]:
            st.session_state.pop(key, None)
        st.rerun()

    st.markdown("---")
    st.markdown("### 📊 Equipe")
    analistas_options = df[[COL_LOGIN, COL_NOME]].drop_duplicates().sort_values(COL_NOME)
    analista_selecionado = st.selectbox(
        "Detalhe individual",
        options=["Todos"] + analistas_options[COL_LOGIN].tolist(),
        format_func=lambda x: "Visão Geral" if x == "Todos" else
            analistas_options[analistas_options[COL_LOGIN]==x][COL_NOME].iloc[0] if len(analistas_options[analistas_options[COL_LOGIN]==x]) > 0 else x,
    )

# Aplicar filtros
df_filtrado = df.copy()

if mes_selecionado != "Todos":
    df_filtrado = df_filtrado[df_filtrado[COL_ANOMES] == mes_selecionado]

if setor_selecionado != "Todos":
    df_filtrado = df_filtrado[df_filtrado["Setor"] == setor_selecionado]


# =====================================================
# KPIs GERAIS
# =====================================================
st.markdown('<div class="section-header">📈 Indicadores Gerais da Equipe</div>', unsafe_allow_html=True)

total_vol = df_filtrado[COL_VOL_TOTAL].sum()
total_dias = df_filtrado.groupby(COL_LOGIN)[COL_DATA].count().sum()
n_analistas = df_filtrado[COL_LOGIN].nunique()
media_diaria_equipe = df_filtrado.groupby(COL_LOGIN)[COL_VOL_TOTAL].sum().mean()

# DPA (filtrando outliers)
dpa_valid = df_filtrado[(df_filtrado[COL_DPA_RESULTADO] >= 0) & (df_filtrado[COL_DPA_RESULTADO] <= 120)]
dpa_media = dpa_valid[COL_DPA_RESULTADO].mean() if not dpa_valid.empty else None

# Período
data_min = df_filtrado[COL_DATA].min()
data_max = df_filtrado[COL_DATA].max()

c1, c2, c3, c4, c5 = st.columns(5)

with c1:
    st.markdown(kpi_card("Volume Total", f"{total_vol:,.0f}", COR_PRIMARIA), unsafe_allow_html=True)
with c2:
    st.markdown(kpi_card("Analistas Ativos", f"{n_analistas}", COR_INFO), unsafe_allow_html=True)
with c3:
    st.markdown(kpi_card("Média/Analista", f"{media_diaria_equipe:,.0f}", COR_SUCESSO), unsafe_allow_html=True)
with c4:
    dpa_display = f"{dpa_media:.0f}" if dpa_media is not None else "—"
    dpa_color = COR_SUCESSO if dpa_media and dpa_media >= 85 else (COR_ALERTA if dpa_media and dpa_media >= 70 else COR_PERIGO)
    st.markdown(kpi_card("DPA Média", dpa_display, dpa_color, suffix="%"), unsafe_allow_html=True)
with c5:
    periodo_str = f"{data_min.strftime('%d/%m')}" if pd.notna(data_min) else "—"
    periodo_str += f" a {data_max.strftime('%d/%m')}" if pd.notna(data_max) else ""
    st.markdown(kpi_card("Período", periodo_str, COR_INFO), unsafe_allow_html=True)


# =====================================================
# VISÃO INDIVIDUAL (se selecionado)
# =====================================================
if analista_selecionado != "Todos":
    df_analista = df_filtrado[df_filtrado[COL_LOGIN] == analista_selecionado]
    if not df_analista.empty:
        nome_analista = df_analista[COL_NOME].iloc[0]
        st.markdown(f'<div class="section-header">👤 Detalhe: {nome_analista}</div>', unsafe_allow_html=True)

        ca1, ca2, ca3, ca4 = st.columns(4)
        vol_ind = df_analista[COL_VOL_TOTAL].sum()
        dias_ind = len(df_analista)
        media_ind = vol_ind / dias_ind if dias_ind > 0 else 0
        dpa_ind_valid = df_analista[(df_analista[COL_DPA_RESULTADO] >= 0) & (df_analista[COL_DPA_RESULTADO] <= 120)]
        dpa_ind = dpa_ind_valid[COL_DPA_RESULTADO].mean() if not dpa_ind_valid.empty else None

        with ca1:
            st.markdown(kpi_card("Volume Total", f"{vol_ind:,.0f}", COR_PRIMARIA), unsafe_allow_html=True)
        with ca2:
            st.markdown(kpi_card("Dias Trabalhados", f"{dias_ind}", COR_INFO), unsafe_allow_html=True)
        with ca3:
            st.markdown(kpi_card("Média Diária", f"{media_ind:.1f}", COR_SUCESSO), unsafe_allow_html=True)
        with ca4:
            dpa_d = f"{dpa_ind:.0f}" if dpa_ind else "—"
            st.markdown(kpi_card("DPA Média", dpa_d, COR_ALERTA, suffix="%"), unsafe_allow_html=True)

        # Gráfico diário do analista
        daily_ind = df_analista.groupby(COL_DATA)[COL_VOL_TOTAL].sum().reset_index()
        daily_ind.columns = ["Data", "Volume"]
        st.bar_chart(daily_ind.set_index("Data"), color=COR_INFO, height=250)

        # Composição de volumes
        vol_breakdown = {}
        for col, label in VOL_COLS.items():
            if col in df_analista.columns:
                v = df_analista[col].sum()
                if v > 0:
                    vol_breakdown[label] = v
        if vol_breakdown:
            comp_df = pd.DataFrame(list(vol_breakdown.items()), columns=["Atividade", "Volume"])
            comp_df = comp_df.sort_values("Volume", ascending=True)
            st.bar_chart(comp_df.set_index("Atividade"), horizontal=True, color=COR_PRIMARIA, height=300)

        st.markdown("---")


# =====================================================
# GRÁFICOS PRINCIPAIS
# =====================================================
tab1, tab2, tab3, tab4 = st.tabs([
    "🏆 Ranking de Produtividade",
    "📅 Evolução Diária",
    "🔍 Composição de Volume",
    "📋 Dados Detalhados",
])

# ---- TAB 1: RANKING ----
with tab1:
    resumo = resumo_geral(df_filtrado)

    if not resumo.empty:
        col_rank1, col_rank2 = st.columns(2)

        with col_rank1:
            st.markdown("#### 📦 Ranking por Volume Total")
            rank_vol = resumo[[COL_LOGIN, COL_NOME, "Setor", COL_VOL_TOTAL, "Dias_Trabalhados", "Media_Diaria"]].copy()
            rank_vol["Nome"] = rank_vol[COL_NOME].apply(primeiro_nome)
            rank_vol = rank_vol.sort_values(COL_VOL_TOTAL, ascending=False).reset_index(drop=True)
            rank_vol.index = rank_vol.index + 1
            rank_vol.index.name = "#"

            display_vol = rank_vol[["Nome", "Setor", COL_VOL_TOTAL, "Dias_Trabalhados", "Media_Diaria"]].copy()
            display_vol.columns = ["Analista", "Setor", "Vol. Total", "Dias", "Média/Dia"]
            st.dataframe(
                display_vol.style.background_gradient(cmap="Blues", subset=["Vol. Total"]),
                use_container_width=True,
                height=500,
            )

        with col_rank2:
            st.markdown("#### ⏱️ Ranking por DPA (Ocupação)")
            rank_dpa = resumo[[COL_LOGIN, COL_NOME, "Setor", "DPA_Media"]].copy()
            rank_dpa["Nome"] = rank_dpa[COL_NOME].apply(primeiro_nome)
            rank_dpa = rank_dpa.dropna(subset=["DPA_Media"])
            rank_dpa = rank_dpa.sort_values("DPA_Media", ascending=False).reset_index(drop=True)
            rank_dpa.index = rank_dpa.index + 1
            rank_dpa.index.name = "#"

            display_dpa = rank_dpa[["Nome", "Setor", "DPA_Media"]].copy()
            display_dpa.columns = ["Analista", "Setor", "DPA %"]
            st.dataframe(
                display_dpa.style
                    .format({"DPA %": "{:.1f}"})
                    .background_gradient(cmap="RdYlGn", subset=["DPA %"], vmin=50, vmax=100),
                use_container_width=True,
                height=500,
            )

        # Ranking por média diária
        st.markdown("#### 📊 Ranking por Média Diária")
        rank_media = resumo[[COL_LOGIN, COL_NOME, "Setor", "Media_Diaria", "Dias_Trabalhados", COL_VOL_TOTAL]].copy()
        rank_media["Nome"] = rank_media[COL_NOME].apply(primeiro_nome)
        rank_media = rank_media.sort_values("Media_Diaria", ascending=False).reset_index(drop=True)
        rank_media.index = rank_media.index + 1
        rank_media.index.name = "#"

        # Gráfico de barras horizontal
        chart_data = rank_media[["Nome", "Media_Diaria"]].set_index("Nome").sort_values("Media_Diaria")
        st.bar_chart(chart_data, horizontal=True, color=COR_PRIMARIA, height=500)

        # ================================================
        # ANÁLISE DETALHADA POR SETOR
        # ================================================
        st.markdown("---")
        st.markdown("#### 📋 Análise Detalhada por Setor")

        sectors_to_show = []
        if setor_selecionado in ("Todos", "RESIDENCIAL"):
            sectors_to_show.append(("RESIDENCIAL", VOL_COLS_RESIDENCIAL, "🏠", "Blues"))
        if setor_selecionado in ("Todos", "EMPRESARIAL"):
            sectors_to_show.append(("EMPRESARIAL", VOL_COLS_EMPRESARIAL, "🏢", "Oranges"))

        for sector_name, sector_specific_vol, sector_icon, sector_cmap in sectors_to_show:
            df_sector = resumo[resumo["Setor"] == sector_name].copy()
            if df_sector.empty:
                continue

            all_vol = {**sector_specific_vol, **VOL_COLS_AMBOS}
            vol_keys = [k for k in all_vol if k in df_sector.columns]

            st.markdown(f"##### {sector_icon} {sector_name}")

            # Build table with comparison metrics
            base = [COL_NOME, COL_VOL_TOTAL, "Dias_Trabalhados", "Media_Diaria", "DPA_Media"]
            base_avail = [c for c in base if c in df_sector.columns]
            detail = df_sector[base_avail + vol_keys].copy()
            detail["Nome"] = detail[COL_NOME].apply(primeiro_nome)

            avg_vol = detail[COL_VOL_TOTAL].mean()
            detail["vs Média"] = ((detail[COL_VOL_TOTAL] / avg_vol - 1) * 100).round(1) if avg_vol > 0 else 0.0

            disp_cols = ["Nome", COL_VOL_TOTAL, "Dias_Trabalhados", "Media_Diaria", "vs Média", "DPA_Media"] + vol_keys
            disp_cols = [c for c in disp_cols if c in detail.columns]
            disp = detail[disp_cols].copy()

            rename = {
                "Nome": "Analista", COL_VOL_TOTAL: "Vol. Total",
                "Dias_Trabalhados": "Dias", "Media_Diaria": "Média/Dia",
                "DPA_Media": "DPA %",
            }
            rename.update({k: all_vol[k] for k in vol_keys})
            disp = disp.rename(columns=rename)
            disp = disp.sort_values("Vol. Total", ascending=False).reset_index(drop=True)
            disp.index += 1
            disp.index.name = "#"

            # Style the table
            fmt = {"DPA %": "{:.1f}", "Média/Dia": "{:.1f}", "vs Média": "{:+.1f}"}
            styled = disp.style.format(fmt, na_rep="—")
            styled = styled.background_gradient(cmap=sector_cmap, subset=["Vol. Total"])
            if disp["vs Média"].notna().any():
                styled = styled.background_gradient(cmap="RdYlGn", subset=["vs Média"], vmin=-50, vmax=50)
            if "DPA %" in disp.columns and disp["DPA %"].notna().any():
                styled = styled.background_gradient(cmap="RdYlGn", subset=["DPA %"], vmin=50, vmax=100)
            for vl in [all_vol[k] for k in vol_keys]:
                if vl in disp.columns and disp[vl].notna().any():
                    styled = styled.background_gradient(cmap=sector_cmap, subset=[vl])

            st.dataframe(styled, use_container_width=True)

            # Best / Worst performer cards
            if len(disp) >= 2:
                best = disp.iloc[0]
                worst = disp.iloc[-1]
                best_dpa_row = None
                if "DPA %" in disp.columns:
                    dpa_valid = disp.dropna(subset=["DPA %"])
                    if not dpa_valid.empty:
                        best_dpa_row = dpa_valid.sort_values("DPA %", ascending=False).iloc[0]

                c_best, c_worst, c_dpa = st.columns(3)
                with c_best:
                    st.markdown(f"""<div class="perf-card perf-best">
                        <strong>🏆 Maior Volume</strong><br>
                        <span style="font-size:1.15rem;font-weight:700;color:#1e8449;">{best['Analista']}</span><br>
                        <span style="font-size:0.82rem;">Vol: {best['Vol. Total']:,.0f} · Média: {best['Média/Dia']:.1f}/dia</span>
                    </div>""", unsafe_allow_html=True)
                with c_worst:
                    st.markdown(f"""<div class="perf-card perf-worst">
                        <strong>⚠️ Menor Volume</strong><br>
                        <span style="font-size:1.15rem;font-weight:700;color:#922b21;">{worst['Analista']}</span><br>
                        <span style="font-size:0.82rem;">Vol: {worst['Vol. Total']:,.0f} · Média: {worst['Média/Dia']:.1f}/dia</span>
                    </div>""", unsafe_allow_html=True)
                with c_dpa:
                    if best_dpa_row is not None:
                        st.markdown(f"""<div class="perf-card perf-dpa">
                            <strong>📊 Melhor DPA</strong><br>
                            <span style="font-size:1.15rem;font-weight:700;color:#1b4f72;">{best_dpa_row['Analista']}</span><br>
                            <span style="font-size:0.82rem;">DPA: {best_dpa_row['DPA %']:.1f}%</span>
                        </div>""", unsafe_allow_html=True)

            st.markdown("")

        # ================================================
        # INSIGHTS POR ANALISTA
        # ================================================
        st.markdown("---")
        st.markdown("#### 💡 Insights — Pontos Fortes e Oportunidades")

        insights_data = []
        for _, row in resumo.sort_values(COL_VOL_TOTAL, ascending=False).iterrows():
            nome = primeiro_nome(row[COL_NOME])
            setor = row["Setor"]
            peers = resumo[resumo["Setor"] == setor]
            n_peers = len(peers)
            if n_peers < 2:
                continue

            if setor == "RESIDENCIAL":
                relevant = {**VOL_COLS_RESIDENCIAL, **VOL_COLS_AMBOS}
            else:
                relevant = {**VOL_COLS_EMPRESARIAL, **VOL_COLS_AMBOS}
            vol_keys_r = [k for k in relevant if k in resumo.columns]

            strengths = []
            weaknesses = []
            for k in vol_keys_r:
                val = row.get(k, 0)
                if pd.isna(val) or val == 0:
                    continue
                rank = int((peers[k].fillna(0) > val).sum() + 1)
                if rank == 1:
                    strengths.append(relevant[k])
                elif rank >= n_peers:
                    weaknesses.append(relevant[k])

            avg_vol = peers[COL_VOL_TOTAL].mean()
            vol_diff = ((row[COL_VOL_TOTAL] / avg_vol - 1) * 100) if avg_vol > 0 else 0
            vol_rank = int((peers[COL_VOL_TOTAL].fillna(0) > row[COL_VOL_TOTAL]).sum() + 1)

            dpa_val = row.get("DPA_Media", None)

            insights_data.append({
                "nome": nome, "setor": setor,
                "vol_total": row[COL_VOL_TOTAL],
                "vol_diff": vol_diff, "vol_rank": vol_rank,
                "dpa": dpa_val,
                "strengths": strengths[:4],
                "weaknesses": weaknesses[:4],
                "n_peers": n_peers,
            })

        col_left, col_right = st.columns(2)
        for i, ins in enumerate(insights_data):
            target_col = col_left if i % 2 == 0 else col_right
            vol_color = "#27AE60" if ins["vol_diff"] >= 0 else "#E74C3C"
            vol_icon = "▲" if ins["vol_diff"] >= 0 else "▼"
            border = "#27AE60" if ins["vol_diff"] >= 10 else ("#E74C3C" if ins["vol_diff"] < -10 else "#2980B9")
            dpa_str = f"{ins['dpa']:.1f}%" if pd.notna(ins["dpa"]) else "—"

            str_tags = "".join(f'<span class="tag-green">{s}</span>' for s in ins["strengths"])
            weak_tags = "".join(f'<span class="tag-red">{w}</span>' for w in ins["weaknesses"])
            if not str_tags:
                str_tags = '<span style="color:#aaa;font-size:0.75rem;">—</span>'
            if not weak_tags:
                weak_tags = '<span style="color:#aaa;font-size:0.75rem;">—</span>'

            with target_col:
                st.markdown(f"""<div class="insight-card" style="border-left-color:{border};">
                    <div style="display:flex;justify-content:space-between;align-items:center;">
                        <div>
                            <strong>{ins['nome']}</strong>
                            <span class="sector-badge">{ins['setor'][:3]}</span>
                            <span class="rank-pill">#{ins['vol_rank']}/{ins['n_peers']}</span>
                        </div>
                        <div style="text-align:right;">
                            <span style="font-weight:700;">{ins['vol_total']:,.0f}</span>
                            <span style="color:{vol_color};font-size:0.82rem;margin-left:3px;">{vol_icon}{abs(ins['vol_diff']):.0f}%</span>
                            <span style="color:#7f8c8d;font-size:0.78rem;margin-left:6px;">DPA:{dpa_str}</span>
                        </div>
                    </div>
                    <div style="margin-top:0.4rem;">
                        <span style="font-size:0.78rem;color:#555;">Forte:</span> {str_tags}
                        <span style="font-size:0.78rem;color:#555;margin-left:8px;">Atenção:</span> {weak_tags}
                    </div>
                </div>""", unsafe_allow_html=True)


# ---- TAB 2: EVOLUÇÃO DIÁRIA ----
with tab2:
    daily = evolucao_diaria(df_filtrado)

    if not daily.empty:
        st.markdown("#### Volume Total da Equipe por Dia")
        chart_daily = daily[[COL_DATA, "Vol_Total"]].set_index(COL_DATA)
        st.area_chart(chart_daily, color=COR_INFO, height=350)

        st.markdown("#### Média por Analista por Dia")
        chart_media = daily[[COL_DATA, "Media_por_Analista"]].set_index(COL_DATA)
        st.line_chart(chart_media, color=COR_SUCESSO, height=300)

        st.markdown("#### Analistas Ativos por Dia")
        chart_an = daily[[COL_DATA, "Analistas"]].set_index(COL_DATA)
        st.bar_chart(chart_an, color=COR_ALERTA, height=250)

    # Comparação mensal (se houver mais de 1 mês)
    meses_unicos = df_filtrado[COL_MES].dropna().unique()
    if len(meses_unicos) > 1 and mes_selecionado == "Todos":
        st.markdown("---")
        st.markdown("#### 📅 Comparação Mensal por Analista")
        mensal = resumo_mensal(df_filtrado)
        if not mensal.empty:
            pivot = mensal.pivot_table(
                index=COL_NOME, columns=COL_MES,
                values="Media_Diaria", aggfunc="first"
            ).reset_index()
            pivot["Nome"] = pivot[COL_NOME].apply(primeiro_nome)
            pivot = pivot.drop(columns=[COL_NOME])
            st.dataframe(pivot, use_container_width=True, hide_index=True)


# ---- TAB 3: COMPOSIÇÃO ----
with tab3:
    comp = composicao_volume(df_filtrado)

    if not comp.empty:
        st.markdown("#### Distribuição por Tipo de Atividade")

        col_pie, col_bar = st.columns([1, 1])

        with col_pie:
            # Top activities
            total = comp["Volume"].sum()
            comp["Percentual"] = (comp["Volume"] / total * 100).round(1)
            comp["Label"] = comp["Atividade"] + " (" + comp["Percentual"].astype(str) + "%)"
            st.dataframe(
                comp[["Atividade", "Volume", "Percentual"]].style.background_gradient(
                    cmap="YlOrRd", subset=["Volume"]
                ),
                use_container_width=True,
                hide_index=True,
            )

        with col_bar:
            chart_comp = comp[["Atividade", "Volume"]].set_index("Atividade").sort_values("Volume")
            st.bar_chart(chart_comp, horizontal=True, color=COR_PRIMARIA, height=450)

    # Composição por setor
    if setor_selecionado == "Todos":
        st.markdown("#### Comparação por Setor")
        for setor in ["EMPRESARIAL", "RESIDENCIAL"]:
            df_setor = df_filtrado[df_filtrado["Setor"] == setor]
            if not df_setor.empty:
                comp_setor = composicao_volume(df_setor)
                if not comp_setor.empty:
                    total_s = comp_setor["Volume"].sum()
                    comp_setor["Percentual"] = (comp_setor["Volume"] / total_s * 100).round(1)
                    st.markdown(f"**{setor}** — Volume Total: {total_s:,.0f}")
                    st.dataframe(
                        comp_setor[["Atividade", "Volume", "Percentual"]],
                        use_container_width=True,
                        hide_index=True,
                        height=200,
                    )


# ---- TAB 4: DADOS DETALHADOS ----
with tab4:
    st.markdown("#### Resumo por Analista")
    resumo_det = resumo_geral(df_filtrado)

    if not resumo_det.empty:
        # Base columns
        display_cols = [COL_NOME, "Setor", "Dias_Trabalhados", COL_VOL_TOTAL, "Media_Diaria", "DPA_Media"]
        display_labels = ["Analista", "Setor", "Dias", "Vol. Total", "Média/Dia", "DPA %"]

        # Add sector-specific volume columns
        sector_vol = get_sector_vol_cols(setor_selecionado, resumo_det.columns)
        vol_keys = list(sector_vol.keys())
        vol_labels = list(sector_vol.values())

        available_base = [c for c in display_cols if c in resumo_det.columns]
        available_labels = display_labels[:len(available_base)]

        det = resumo_det[available_base + vol_keys].copy()
        det.columns = available_labels + vol_labels
        det = det.sort_values("Vol. Total", ascending=False).reset_index(drop=True)
        det.index = det.index + 1
        det.index.name = "#"

        format_dict = {}
        if "DPA %" in det.columns:
            format_dict["DPA %"] = "{:.1f}"
        st.dataframe(
            det.style.format(format_dict) if format_dict else det,
            use_container_width=True,
        )

    st.markdown("---")
    st.markdown("#### Dados Brutos (Filtrados)")
    cols_to_show = [COL_LOGIN, COL_NOME, "Setor", COL_DATA, COL_MES, COL_VOL_TOTAL, COL_DPA_RESULTADO]
    vol_cols_existing = [c for c in VOL_COLS.keys() if c in df_filtrado.columns]
    cols_to_show += vol_cols_existing
    cols_existing = [c for c in cols_to_show if c in df_filtrado.columns]

    st.dataframe(
        df_filtrado[cols_existing].sort_values([COL_NOME, COL_DATA]),
        use_container_width=True,
        height=500,
    )

    # Download
    csv = df_filtrado[cols_existing].to_csv(index=False).encode("utf-8")
    st.download_button(
        "📥 Baixar dados filtrados (CSV)",
        csv,
        "produtividade_equipe.csv",
        "text/csv",
    )


# =====================================================
# FOOTER
# =====================================================
st.markdown("---")
st.caption(
    f"Dashboard de Produtividade COP Rede · "
    f"{len(df_filtrado)} registros carregados · "
    f"{n_analistas} analistas · "
    f"Dados de {data_min.strftime('%d/%m/%Y') if pd.notna(data_min) else '?'} "
    f"a {data_max.strftime('%d/%m/%Y') if pd.notna(data_max) else '?'}"
)
