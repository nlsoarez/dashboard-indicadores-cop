import traceback
import pandas as pd
import numpy as np
import streamlit as st

from src.config import (
    BASE_EQUIPE, EQUIPE_IDS, VOL_COLS,
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
# UPLOAD
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

if uploaded_file is None:
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
        df = load_produtividade(uploaded_file)

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
                display_dpa.style.background_gradient(cmap="RdYlGn", subset=["DPA %"], vmin=50, vmax=100),
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
        display_cols = [COL_NOME, "Setor", "Dias_Trabalhados", COL_VOL_TOTAL, "Media_Diaria", "DPA_Media"]
        available_cols = [c for c in display_cols if c in resumo_det.columns]
        det = resumo_det[available_cols].copy()
        det.columns = ["Analista", "Setor", "Dias", "Vol. Total", "Média/Dia", "DPA %"][:len(available_cols)]
        det = det.sort_values("Vol. Total", ascending=False).reset_index(drop=True)
        det.index = det.index + 1
        det.index.name = "#"
        st.dataframe(det, use_container_width=True)

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
