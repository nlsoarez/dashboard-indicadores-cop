import traceback
import io
import pandas as pd
import numpy as np
import streamlit as st

from src.config import (
    BASE_EQUIPE, EQUIPE_IDS, LIDERES_IDS, VOL_COLS,
    VOL_COLS_RESIDENCIAL, VOL_COLS_EMPRESARIAL, VOL_COLS_AMBOS,
    COL_LOGIN, COL_NOME, COL_BASE, COL_DATA, COL_MES, COL_ANOMES,
    COL_VOL_TOTAL, COL_DPA_RESULTADO,
    COR_PRIMARIA, COR_SUCESSO, COR_ALERTA, COR_PERIGO, COR_INFO,
    # ETIT
    ETIT_COL_LOGIN, ETIT_COL_DEMANDA, ETIT_COL_VOLUME,
    ETIT_COL_STATUS, ETIT_COL_TIPO, ETIT_COL_CAUSA,
    ETIT_COL_REGIONAL, ETIT_COL_TURNO, ETIT_COL_TMA, ETIT_COL_TMR,
    ETIT_COL_DT_ACIONAMENTO, ETIT_COL_ANOMES, ETIT_COL_INDICADOR_VAL,
    ETIT_COL_NOTA, ETIT_COL_AREA, ETIT_COL_CIDADE, ETIT_COL_UF,
)
from src.processors import (
    load_produtividade, resumo_mensal, resumo_geral,
    evolucao_diaria, composicao_volume, primeiro_nome,
    # ETIT
    load_etit, etit_resumo_analista, etit_por_demanda,
    etit_por_tipo, etit_por_causa, etit_por_regional,
    etit_por_turno, etit_evolucao_diaria,
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
# CSS — tema escuro profissional
# =====================================================
st.markdown("""
<style>
    /* Header */
    .main-header {
        background: linear-gradient(135deg, #1B4F72 0%, #2980B9 100%);
        padding: 1.5rem 2rem;
        border-radius: 12px;
        margin-bottom: 1.5rem;
    }
    .main-header h1 { color: #fff !important; margin: 0; font-size: 1.8rem; }
    .main-header p  { color: #D6EAF8 !important; margin: 0.3rem 0 0 0; font-size: 0.95rem; }

    /* KPI Cards */
    .kpi-card {
        background: var(--secondary-background-color);
        border-radius: 10px;
        padding: 1.2rem;
        border-left: 4px solid;
        text-align: center;
    }
    .kpi-card .kpi-value { font-size: 1.8rem; font-weight: 700; margin: 0.3rem 0; }
    .kpi-card .kpi-label {
        font-size: 0.78rem; opacity: 0.55;
        text-transform: uppercase; letter-spacing: 0.5px;
    }
    .kpi-card .kpi-delta { font-size: 0.8rem; margin-top: 0.2rem; }

    /* Section headers */
    .section-header {
        font-size: 1.15rem; font-weight: 600;
        color: #5DADE2;
        border-bottom: 2px solid rgba(41,128,185,0.4);
        padding-bottom: 0.4rem;
        margin: 1.5rem 0 1rem 0;
    }

    /* Performance highlight cards */
    .perf-card {
        padding: 0.9rem 1rem; border-radius: 10px;
        border-left: 4px solid; margin-bottom: 0.5rem;
    }
    .perf-best  { background: rgba(39,174,96,0.12);  border-left-color: #2ECC71; }
    .perf-worst { background: rgba(231,76,60,0.12);   border-left-color: #E74C3C; }
    .perf-dpa   { background: rgba(41,128,185,0.12);  border-left-color: #5DADE2; }
    .perf-card .p-title  { font-size: 0.8rem; font-weight: 600; margin-bottom: 0.25rem; }
    .perf-card .p-name   { font-size: 1.1rem; font-weight: 700; }
    .perf-card .p-detail { font-size: 0.82rem; opacity: 0.7; margin-top: 0.15rem; }

    /* Insight cards */
    .insight-card {
        background: var(--secondary-background-color);
        border-radius: 10px;
        padding: 0.85rem 1rem;
        margin-bottom: 0.6rem;
        border-left: 4px solid #5DADE2;
    }
    .tag-green {
        background: rgba(46,204,113,0.18); color: #2ECC71;
        padding: 2px 8px; border-radius: 10px;
        font-size: 0.73rem; font-weight: 500;
        display: inline-block; margin: 1px 2px;
    }
    .tag-red {
        background: rgba(231,76,60,0.18); color: #E74C3C;
        padding: 2px 8px; border-radius: 10px;
        font-size: 0.73rem; font-weight: 500;
        display: inline-block; margin: 1px 2px;
    }
    .sector-badge {
        background: rgba(93,173,226,0.15); color: #5DADE2;
        padding: 2px 8px; border-radius: 8px;
        font-size: 0.72rem; font-weight: 500; margin-left: 6px;
    }
    .rank-pill {
        background: rgba(255,255,255,0.06);
        padding: 2px 8px; border-radius: 8px;
        font-size: 0.72rem; opacity: 0.55; margin-left: 3px;
    }

    /* Leader cards */
    .leader-card {
        background: var(--secondary-background-color);
        border-radius: 12px;
        padding: 1rem 1.2rem;
        border-top: 3px solid #F1C40F;
        margin-bottom: 0.8rem;
    }
    .leader-card .l-name  { font-size: 1.05rem; font-weight: 700; }
    .leader-card .l-badge {
        background: rgba(241,196,15,0.15); color: #F1C40F;
        padding: 2px 8px; border-radius: 8px;
        font-size: 0.72rem; font-weight: 600; margin-left: 6px;
    }
    .leader-card .l-stat  { font-size: 0.82rem; opacity: 0.7; margin-top: 0.3rem; }
    .leader-card .l-vol   { font-size: 1.4rem; font-weight: 700; color: #5DADE2; }

    /* Sector header pills */
    .sector-header {
        display: inline-block;
        padding: 0.35rem 1rem; border-radius: 8px;
        font-weight: 600; font-size: 0.9rem;
        margin-bottom: 0.6rem;
    }
    .sector-res { background: rgba(41,128,185,0.15); color: #5DADE2; }
    .sector-emp { background: rgba(243,156,18,0.15); color: #F39C12; }

    /* ETIT highlight card */
    .etit-card {
        background: var(--secondary-background-color);
        border-radius: 10px;
        padding: 1rem 1.2rem;
        border-left: 4px solid #8E44AD;
        margin-bottom: 0.6rem;
    }

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


def get_sector_vol_cols(setor, available_cols):
    """Returns {col_name: display_label} for the given sector filter."""
    cols = {}
    if setor in ("Todos", "RESIDENCIAL"):
        cols.update(VOL_COLS_RESIDENCIAL)
    if setor in ("Todos", "EMPRESARIAL"):
        cols.update(VOL_COLS_EMPRESARIAL)
    cols.update(VOL_COLS_AMBOS)
    return {k: v for k, v in cols.items() if k in available_cols}


def build_insights(resumo_df, setor_filter):
    """Computes strengths/weaknesses for each analyst vs sector peers."""
    data = []
    for _, row in resumo_df.sort_values(COL_VOL_TOTAL, ascending=False).iterrows():
        nome = primeiro_nome(row[COL_NOME])
        setor = row["Setor"]
        peers = resumo_df[resumo_df["Setor"] == setor]
        n_peers = len(peers)
        if n_peers < 2:
            continue

        if setor == "RESIDENCIAL":
            relevant = {**VOL_COLS_RESIDENCIAL, **VOL_COLS_AMBOS}
        else:
            relevant = {**VOL_COLS_EMPRESARIAL, **VOL_COLS_AMBOS}
        vol_keys_r = [k for k in relevant if k in resumo_df.columns]

        strengths, weaknesses = [], []
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

        data.append({
            "nome": nome, "setor": setor, "login": row[COL_LOGIN],
            "vol_total": row[COL_VOL_TOTAL],
            "media_diaria": row.get("Media_Diaria", 0),
            "dias": row.get("Dias_Trabalhados", 0),
            "vol_diff": vol_diff, "vol_rank": vol_rank,
            "dpa": dpa_val,
            "strengths": strengths[:4], "weaknesses": weaknesses[:4],
            "n_peers": n_peers,
        })
    return data


def render_insight_cards(insights_list):
    """Renders insight cards in 2-column layout."""
    col_l, col_r = st.columns(2)
    for i, ins in enumerate(insights_list):
        target = col_l if i % 2 == 0 else col_r
        vol_color = "#2ECC71" if ins["vol_diff"] >= 0 else "#E74C3C"
        vol_icon = "▲" if ins["vol_diff"] >= 0 else "▼"
        border = "#2ECC71" if ins["vol_diff"] >= 10 else ("#E74C3C" if ins["vol_diff"] < -10 else "#5DADE2")
        dpa_str = f"{ins['dpa']:.1f}%" if pd.notna(ins["dpa"]) else "—"

        str_tags = "".join(f'<span class="tag-green">{s}</span>' for s in ins["strengths"])
        weak_tags = "".join(f'<span class="tag-red">{w}</span>' for w in ins["weaknesses"])
        if not str_tags:
            str_tags = '<span style="opacity:0.35;font-size:0.73rem;">—</span>'
        if not weak_tags:
            weak_tags = '<span style="opacity:0.35;font-size:0.73rem;">—</span>'

        with target:
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
                        <span style="opacity:0.5;font-size:0.78rem;margin-left:6px;">DPA:{dpa_str}</span>
                    </div>
                </div>
                <div style="margin-top:0.4rem;">
                    <span style="font-size:0.78rem;opacity:0.5;">Forte:</span> {str_tags}
                    <span style="font-size:0.78rem;opacity:0.5;margin-left:8px;">Atenção:</span> {weak_tags}
                </div>
            </div>""", unsafe_allow_html=True)


def render_sector_table(resumo_df, sector_name, sector_vol, sector_cmap):
    """Renders a styled sector detail table with best/worst cards."""
    df_sec = resumo_df[resumo_df["Setor"] == sector_name].copy()
    if df_sec.empty:
        return

    all_vol = {**sector_vol, **VOL_COLS_AMBOS}
    vol_keys = [k for k in all_vol if k in df_sec.columns]

    css_cls = "sector-res" if sector_name == "RESIDENCIAL" else "sector-emp"
    icon = "🏠" if sector_name == "RESIDENCIAL" else "🏢"
    st.markdown(f'<span class="sector-header {css_cls}">{icon} {sector_name}</span>', unsafe_allow_html=True)

    base = [COL_NOME, COL_VOL_TOTAL, "Dias_Trabalhados", "Media_Diaria", "DPA_Media"]
    base_avail = [c for c in base if c in df_sec.columns]
    detail = df_sec[base_avail + vol_keys].copy()
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

    # Best / Worst cards
    if len(disp) >= 2:
        best = disp.iloc[0]
        worst = disp.iloc[-1]
        best_dpa_row = None
        if "DPA %" in disp.columns:
            dpa_v = disp.dropna(subset=["DPA %"])
            if not dpa_v.empty:
                best_dpa_row = dpa_v.sort_values("DPA %", ascending=False).iloc[0]

        c1, c2, c3 = st.columns(3)
        with c1:
            st.markdown(f"""<div class="perf-card perf-best">
                <div class="p-title">🏆 Maior Volume</div>
                <div class="p-name" style="color:#2ECC71;">{best['Analista']}</div>
                <div class="p-detail">Vol: {best['Vol. Total']:,.0f} · Média: {best['Média/Dia']:.1f}/dia</div>
            </div>""", unsafe_allow_html=True)
        with c2:
            st.markdown(f"""<div class="perf-card perf-worst">
                <div class="p-title">⚠️ Menor Volume</div>
                <div class="p-name" style="color:#E74C3C;">{worst['Analista']}</div>
                <div class="p-detail">Vol: {worst['Vol. Total']:,.0f} · Média: {worst['Média/Dia']:.1f}/dia</div>
            </div>""", unsafe_allow_html=True)
        with c3:
            if best_dpa_row is not None:
                st.markdown(f"""<div class="perf-card perf-dpa">
                    <div class="p-title">📊 Melhor DPA</div>
                    <div class="p-name" style="color:#5DADE2;">{best_dpa_row['Analista']}</div>
                    <div class="p-detail">DPA: {best_dpa_row['DPA %']:.1f}%</div>
                </div>""", unsafe_allow_html=True)

    st.markdown("")


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
    col_upload1, col_upload2, col_info = st.columns([2, 2, 1])
    with col_upload1:
        uploaded_file = st.file_uploader(
            "📁 Planilha de Produtividade (Analítico)",
            type=["xlsx", "xls"],
            help="Planilha com aba 'Analítico Produtividade 2026' ou similar",
            key="upload_prod",
        )
    with col_upload2:
        uploaded_etit = st.file_uploader(
            "📁 Planilha Analítico Empresarial (ETIT)",
            type=["xlsx", "xls"],
            help="Planilha com dados ETIT POR EVENTO — opcional",
            key="upload_etit",
        )
    with col_info:
        st.info(
            f"**Equipe monitorada:** {len(EQUIPE_IDS)} analistas\n\n"
            f"Empresarial: {len(BASE_EQUIPE[BASE_EQUIPE['Setor']=='EMPRESARIAL'])} · "
            f"Residencial: {len(BASE_EQUIPE[BASE_EQUIPE['Setor']=='RESIDENCIAL'])}"
        )

if uploaded_file is not None:
    st.session_state["uploaded_bytes"] = uploaded_file.getvalue()
    st.session_state["uploaded_name"] = uploaded_file.name

if uploaded_etit is not None:
    st.session_state["uploaded_etit_bytes"] = uploaded_etit.getvalue()
    st.session_state["uploaded_etit_name"] = uploaded_etit.name

if "uploaded_bytes" not in st.session_state:
    st.markdown("---")
    st.markdown("### 👋 Bem-vindo!")
    st.markdown(
        "Faça upload da planilha **Produtividade COP Rede - Analítico** acima para "
        "visualizar os dados de produtividade da sua equipe.\n\n"
        "Opcionalmente, faça upload da planilha **Analítico Empresarial** para "
        "visualizar os dados de **ETIT POR EVENTO**."
    )
    with st.expander("📋 Analistas monitorados"):
        st.dataframe(BASE_EQUIPE, use_container_width=True, hide_index=True)
    st.stop()


# =====================================================
# PROCESSAR DADOS — Produtividade
# =====================================================
try:
    with st.spinner("Carregando e processando dados de produtividade..."):
        file_obj = io.BytesIO(st.session_state["uploaded_bytes"])
        df = load_produtividade(file_obj)

    if df.empty:
        st.error("Nenhum analista da equipe encontrado na planilha de produtividade. Verifique o arquivo.")
        st.stop()

except Exception as e:
    st.error(f"Erro ao processar a planilha de produtividade: {e}")
    with st.expander("Detalhes do erro"):
        st.code(traceback.format_exc())
    st.stop()


# =====================================================
# PROCESSAR DADOS — ETIT (opcional)
# =====================================================
df_etit = pd.DataFrame()
etit_loaded = False

if "uploaded_etit_bytes" in st.session_state:
    try:
        with st.spinner("Carregando dados ETIT POR EVENTO..."):
            etit_obj = io.BytesIO(st.session_state["uploaded_etit_bytes"])
            df_etit = load_etit(etit_obj)
            etit_loaded = not df_etit.empty
            if not etit_loaded:
                st.warning("Nenhum analista da equipe encontrado nos dados ETIT POR EVENTO.")
    except Exception as e:
        st.warning(f"Erro ao processar planilha ETIT: {e}")
        with st.expander("Detalhes do erro"):
            st.code(traceback.format_exc())


# =====================================================
# SIDEBAR - FILTROS
# =====================================================
with st.sidebar:
    st.markdown("### 🔧 Filtros")

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
        for key in ["uploaded_bytes", "uploaded_name", "uploaded_etit_bytes", "uploaded_etit_name"]:
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

    # ETIT status
    if etit_loaded:
        st.markdown("---")
        st.success(f"✅ ETIT: {len(df_etit)} eventos carregados")

# Aplicar filtros — Produtividade
df_filtrado = df.copy()
if mes_selecionado != "Todos":
    df_filtrado = df_filtrado[df_filtrado[COL_ANOMES] == mes_selecionado]
if setor_selecionado != "Todos":
    df_filtrado = df_filtrado[df_filtrado["Setor"] == setor_selecionado]

# Aplicar filtros — ETIT
df_etit_filtrado = df_etit.copy()
if etit_loaded:
    if mes_selecionado != "Todos" and ETIT_COL_ANOMES in df_etit_filtrado.columns:
        df_etit_filtrado = df_etit_filtrado[df_etit_filtrado[ETIT_COL_ANOMES] == str(mes_selecionado)]
    if setor_selecionado != "Todos":
        df_etit_filtrado = df_etit_filtrado[df_etit_filtrado["Setor"] == setor_selecionado]
    if analista_selecionado != "Todos":
        df_etit_filtrado = df_etit_filtrado[df_etit_filtrado[ETIT_COL_LOGIN] == analista_selecionado]


# =====================================================
# KPIs GERAIS
# =====================================================
st.markdown('<div class="section-header">📈 Indicadores Gerais da Equipe</div>', unsafe_allow_html=True)

total_vol = df_filtrado[COL_VOL_TOTAL].sum()
n_analistas = df_filtrado[COL_LOGIN].nunique()
media_diaria_equipe = df_filtrado.groupby(COL_LOGIN)[COL_VOL_TOTAL].sum().mean()
dpa_valid = df_filtrado[(df_filtrado[COL_DPA_RESULTADO] >= 0) & (df_filtrado[COL_DPA_RESULTADO] <= 120)]
dpa_media = dpa_valid[COL_DPA_RESULTADO].mean() if not dpa_valid.empty else None
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

        daily_ind = df_analista.groupby(COL_DATA)[COL_VOL_TOTAL].sum().reset_index()
        daily_ind.columns = ["Data", "Volume"]
        st.bar_chart(daily_ind.set_index("Data"), color=COR_INFO, height=250)

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

        # ETIT individual detail
        if etit_loaded and not df_etit_filtrado.empty:
            st.markdown("##### ⚡ ETIT POR EVENTO — Este Analista")
            etit_ind = df_etit_filtrado[df_etit_filtrado[ETIT_COL_LOGIN] == analista_selecionado]
            if not etit_ind.empty:
                ei1, ei2, ei3, ei4 = st.columns(4)
                etit_total = etit_ind[ETIT_COL_VOLUME].sum()
                etit_ader = etit_ind[ETIT_COL_INDICADOR_VAL].sum()
                etit_pct = (etit_ader / etit_total * 100) if etit_total > 0 else 0
                etit_tma = etit_ind[ETIT_COL_TMA].mean()
                with ei1:
                    st.markdown(kpi_card("Eventos ETIT", f"{etit_total:,.0f}", "#8E44AD"), unsafe_allow_html=True)
                with ei2:
                    st.markdown(kpi_card("Aderentes", f"{etit_ader:,.0f}", COR_SUCESSO), unsafe_allow_html=True)
                with ei3:
                    ad_color = COR_SUCESSO if etit_pct >= 90 else (COR_ALERTA if etit_pct >= 70 else COR_PERIGO)
                    st.markdown(kpi_card("Aderência", f"{etit_pct:.1f}", ad_color, suffix="%"), unsafe_allow_html=True)
                with ei4:
                    st.markdown(kpi_card("TMA Médio", f"{etit_tma:.4f}", COR_INFO), unsafe_allow_html=True)
            else:
                st.caption("Nenhum evento ETIT encontrado para este analista no período.")

        st.markdown("---")


# =====================================================
# TABS PRINCIPAIS
# =====================================================
tab_labels = [
    "🏆 Ranking",
    "👑 Líderes",
    "📅 Evolução Diária",
    "🔍 Composição",
    "📋 Dados Detalhados",
]
if etit_loaded:
    tab_labels.append("⚡ ETIT por Evento")

tabs = st.tabs(tab_labels)

# ---- TAB 1: RANKING ----
with tabs[0]:
    resumo = resumo_geral(df_filtrado)

    if not resumo.empty:
        # Exclude leaders from team ranking
        resumo_equipe = resumo[~resumo[COL_LOGIN].isin(LIDERES_IDS)].copy()

        col_rank1, col_rank2 = st.columns(2)

        with col_rank1:
            st.markdown("#### 📦 Ranking por Volume Total")
            rank_vol = resumo_equipe[[COL_LOGIN, COL_NOME, "Setor", COL_VOL_TOTAL, "Dias_Trabalhados", "Media_Diaria"]].copy()
            rank_vol["Nome"] = rank_vol[COL_NOME].apply(primeiro_nome)
            rank_vol = rank_vol.sort_values(COL_VOL_TOTAL, ascending=False).reset_index(drop=True)
            rank_vol.index = rank_vol.index + 1
            rank_vol.index.name = "#"
            display_vol = rank_vol[["Nome", "Setor", COL_VOL_TOTAL, "Dias_Trabalhados", "Media_Diaria"]].copy()
            display_vol.columns = ["Analista", "Setor", "Vol. Total", "Dias", "Média/Dia"]
            st.dataframe(
                display_vol.style.background_gradient(cmap="Blues", subset=["Vol. Total"]),
                use_container_width=True, height=500,
            )

        with col_rank2:
            st.markdown("#### ⏱️ Ranking por DPA (Ocupação)")
            rank_dpa = resumo_equipe[[COL_LOGIN, COL_NOME, "Setor", "DPA_Media"]].copy()
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
                use_container_width=True, height=500,
            )

        # Bar chart
        st.markdown("#### 📊 Ranking por Média Diária")
        rank_media = resumo_equipe[[COL_NOME, "Media_Diaria"]].copy()
        rank_media["Nome"] = rank_media[COL_NOME].apply(primeiro_nome)
        chart_data = rank_media[["Nome", "Media_Diaria"]].set_index("Nome").sort_values("Media_Diaria")
        st.bar_chart(chart_data, horizontal=True, color=COR_PRIMARIA, height=500)

        # ============ SECTOR DETAIL ============
        st.markdown("---")
        st.markdown("#### 📋 Análise Detalhada por Setor")

        if setor_selecionado in ("Todos", "RESIDENCIAL"):
            render_sector_table(resumo_equipe, "RESIDENCIAL", VOL_COLS_RESIDENCIAL, "Blues")
        if setor_selecionado in ("Todos", "EMPRESARIAL"):
            render_sector_table(resumo_equipe, "EMPRESARIAL", VOL_COLS_EMPRESARIAL, "Oranges")

        # ============ INSIGHTS ============
        st.markdown("---")
        st.markdown("#### 💡 Insights — Pontos Fortes e Oportunidades")
        insights = build_insights(resumo_equipe, setor_selecionado)
        render_insight_cards(insights)


# ---- TAB LÍDERES ----
with tabs[1]:
    resumo_full = resumo_geral(df_filtrado)
    resumo_lid = resumo_full[resumo_full[COL_LOGIN].isin(LIDERES_IDS)].copy()
    resumo_equipe_all = resumo_full[~resumo_full[COL_LOGIN].isin(LIDERES_IDS)].copy()

    if not resumo_lid.empty:
        st.markdown("#### 👑 Visão dos Líderes")
        st.caption("Comparação entre os líderes e as médias das suas equipes.")

        # Leader cards
        cols_lid = st.columns(len(resumo_lid))
        for idx, (_, lrow) in enumerate(resumo_lid.sort_values(COL_VOL_TOTAL, ascending=False).iterrows()):
            nome_l = primeiro_nome(lrow[COL_NOME])
            setor_l = lrow["Setor"]
            vol_l = lrow[COL_VOL_TOTAL]
            media_l = lrow.get("Media_Diaria", 0)
            dias_l = lrow.get("Dias_Trabalhados", 0)
            dpa_l = lrow.get("DPA_Media", None)
            dpa_l_str = f"{dpa_l:.1f}%" if pd.notna(dpa_l) else "—"

            # Compare vs team average
            team = resumo_equipe_all[resumo_equipe_all["Setor"] == setor_l]
            team_avg_vol = team[COL_VOL_TOTAL].mean() if not team.empty else 0
            team_avg_media = team["Media_Diaria"].mean() if not team.empty else 0
            diff_vol = ((vol_l / team_avg_vol - 1) * 100) if team_avg_vol > 0 else 0
            diff_media = ((media_l / team_avg_media - 1) * 100) if team_avg_media > 0 else 0

            badge_setor = "RES" if setor_l == "RESIDENCIAL" else "EMP"
            diff_vol_color = "#2ECC71" if diff_vol >= 0 else "#E74C3C"
            diff_vol_icon = "▲" if diff_vol >= 0 else "▼"

            with cols_lid[idx]:
                st.markdown(f"""<div class="leader-card">
                    <div style="display:flex;justify-content:space-between;align-items:center;">
                        <div>
                            <span class="l-name">{nome_l}</span>
                            <span class="l-badge">{badge_setor}</span>
                        </div>
                    </div>
                    <div class="l-vol">{vol_l:,.0f}</div>
                    <div class="l-stat">Média: {media_l:.1f}/dia · Dias: {dias_l:.0f} · DPA: {dpa_l_str}</div>
                    <div class="l-stat" style="margin-top:0.2rem;">
                        vs Equipe: <span style="color:{diff_vol_color};font-weight:600;">{diff_vol_icon}{abs(diff_vol):.0f}% vol</span>
                        · <span style="color:{diff_vol_color};font-weight:600;">{diff_vol_icon}{abs(diff_media):.0f}% média</span>
                    </div>
                </div>""", unsafe_allow_html=True)

        # Leaders detailed table
        st.markdown("---")
        st.markdown("#### 📊 Comparação Detalhada")

        sector_vol_lid = get_sector_vol_cols(setor_selecionado, resumo_lid.columns)
        vol_keys_lid = list(sector_vol_lid.keys())
        base_cols_lid = [COL_NOME, "Setor", COL_VOL_TOTAL, "Dias_Trabalhados", "Media_Diaria", "DPA_Media"]
        avail_lid = [c for c in base_cols_lid if c in resumo_lid.columns]
        det_lid = resumo_lid[avail_lid + vol_keys_lid].copy()
        det_lid[COL_NOME] = det_lid[COL_NOME].apply(primeiro_nome)
        rename_lid = {
            COL_NOME: "Líder", "Setor": "Setor", COL_VOL_TOTAL: "Vol. Total",
            "Dias_Trabalhados": "Dias", "Media_Diaria": "Média/Dia", "DPA_Media": "DPA %",
        }
        rename_lid.update(sector_vol_lid)
        det_lid = det_lid.rename(columns=rename_lid)
        det_lid = det_lid.sort_values("Vol. Total", ascending=False).reset_index(drop=True)
        det_lid.index += 1
        det_lid.index.name = "#"

        fmt_lid = {"DPA %": "{:.1f}", "Média/Dia": "{:.1f}"}
        styled_lid = det_lid.style.format(fmt_lid, na_rep="—")
        styled_lid = styled_lid.background_gradient(cmap="YlOrBr", subset=["Vol. Total"])
        if "DPA %" in det_lid.columns and det_lid["DPA %"].notna().any():
            styled_lid = styled_lid.background_gradient(cmap="RdYlGn", subset=["DPA %"], vmin=50, vmax=100)
        vol_lbl_lid = [sector_vol_lid[k] for k in vol_keys_lid if sector_vol_lid[k] in det_lid.columns]
        for vl in vol_lbl_lid:
            if det_lid[vl].notna().any():
                styled_lid = styled_lid.background_gradient(cmap="YlOrBr", subset=[vl])
        st.dataframe(styled_lid, use_container_width=True)

        # Leaders vs team comparison
        st.markdown("---")
        st.markdown("#### 📈 Líderes vs Média da Equipe")

        for setor_comp in ["RESIDENCIAL", "EMPRESARIAL"]:
            lids_setor = resumo_lid[resumo_lid["Setor"] == setor_comp]
            team_setor = resumo_equipe_all[resumo_equipe_all["Setor"] == setor_comp]
            if lids_setor.empty:
                continue

            css_s = "sector-res" if setor_comp == "RESIDENCIAL" else "sector-emp"
            icon_s = "🏠" if setor_comp == "RESIDENCIAL" else "🏢"
            st.markdown(f'<span class="sector-header {css_s}">{icon_s} {setor_comp}</span>', unsafe_allow_html=True)

            team_avg_vol_s = team_setor[COL_VOL_TOTAL].mean() if not team_setor.empty else 0
            team_avg_media_s = team_setor["Media_Diaria"].mean() if not team_setor.empty else 0
            team_avg_dpa_s = team_setor["DPA_Media"].mean() if not team_setor.empty and team_setor["DPA_Media"].notna().any() else None

            rows = []
            for _, lr in lids_setor.iterrows():
                nome_r = primeiro_nome(lr[COL_NOME])
                rows.append({
                    "Analista": nome_r,
                    "Vol. Total": lr[COL_VOL_TOTAL],
                    "Média/Dia": lr.get("Media_Diaria", 0),
                    "DPA %": lr.get("DPA_Media", None),
                })
            rows.append({
                "Analista": "📊 MÉDIA EQUIPE",
                "Vol. Total": team_avg_vol_s,
                "Média/Dia": team_avg_media_s,
                "DPA %": team_avg_dpa_s,
            })
            comp_df = pd.DataFrame(rows)
            comp_df = comp_df.set_index("Analista")
            st.dataframe(
                comp_df.style
                    .format({"DPA %": "{:.1f}", "Média/Dia": "{:.1f}", "Vol. Total": "{:,.0f}"}, na_rep="—"),
                use_container_width=True,
            )
            st.markdown("")

        # Insights for leaders
        st.markdown("---")
        st.markdown("#### 💡 Insights dos Líderes")
        lid_insights = build_insights(resumo_full, setor_selecionado)
        lid_insights = [i for i in lid_insights if i["login"] in LIDERES_IDS]
        render_insight_cards(lid_insights)

    else:
        st.info("Nenhum líder encontrado nos dados filtrados.")


# ---- TAB 2: EVOLUÇÃO DIÁRIA ----
with tabs[2]:
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
with tabs[3]:
    comp = composicao_volume(df_filtrado)

    if not comp.empty:
        st.markdown("#### Distribuição por Tipo de Atividade")
        col_pie, col_bar = st.columns([1, 1])

        with col_pie:
            total = comp["Volume"].sum()
            comp["Percentual"] = (comp["Volume"] / total * 100).round(1)
            st.dataframe(
                comp[["Atividade", "Volume", "Percentual"]].style.background_gradient(
                    cmap="YlOrRd", subset=["Volume"]
                ),
                use_container_width=True, hide_index=True,
            )

        with col_bar:
            chart_comp = comp[["Atividade", "Volume"]].set_index("Atividade").sort_values("Volume")
            st.bar_chart(chart_comp, horizontal=True, color=COR_PRIMARIA, height=450)

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
                        use_container_width=True, hide_index=True, height=200,
                    )


# ---- TAB 4: DADOS DETALHADOS ----
with tabs[4]:
    st.markdown("#### Resumo por Analista")
    resumo_det = resumo_geral(df_filtrado)

    if not resumo_det.empty:
        display_cols = [COL_NOME, "Setor", "Dias_Trabalhados", COL_VOL_TOTAL, "Media_Diaria", "DPA_Media"]
        display_labels = ["Analista", "Setor", "Dias", "Vol. Total", "Média/Dia", "DPA %"]

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
        use_container_width=True, height=500,
    )

    csv = df_filtrado[cols_existing].to_csv(index=False).encode("utf-8")
    st.download_button(
        "📥 Baixar dados filtrados (CSV)",
        csv, "produtividade_equipe.csv", "text/csv",
    )


# ---- TAB 5: ETIT POR EVENTO ----
if etit_loaded:
    with tabs[5]:
        st.markdown("#### ⚡ ETIT POR EVENTO — Análise da Equipe")

        if df_etit_filtrado.empty:
            st.warning("Nenhum dado ETIT POR EVENTO encontrado com os filtros atuais.")
        else:
            # ---- KPIs ETIT ----
            etit_total_eventos = df_etit_filtrado[ETIT_COL_VOLUME].sum()
            etit_total_ader = df_etit_filtrado[ETIT_COL_INDICADOR_VAL].sum()
            etit_pct_ader = (etit_total_ader / etit_total_eventos * 100) if etit_total_eventos > 0 else 0
            etit_n_analistas = df_etit_filtrado[ETIT_COL_LOGIN].nunique()
            etit_tma_geral = df_etit_filtrado[ETIT_COL_TMA].mean()
            etit_tmr_geral = df_etit_filtrado[ETIT_COL_TMR].mean()

            ek1, ek2, ek3, ek4, ek5, ek6 = st.columns(6)
            with ek1:
                st.markdown(kpi_card("Total Eventos", f"{etit_total_eventos:,.0f}", "#8E44AD"), unsafe_allow_html=True)
            with ek2:
                st.markdown(kpi_card("Aderentes", f"{etit_total_ader:,.0f}", COR_SUCESSO), unsafe_allow_html=True)
            with ek3:
                ad_c = COR_SUCESSO if etit_pct_ader >= 90 else (COR_ALERTA if etit_pct_ader >= 70 else COR_PERIGO)
                st.markdown(kpi_card("Aderência", f"{etit_pct_ader:.1f}", ad_c, suffix="%"), unsafe_allow_html=True)
            with ek4:
                st.markdown(kpi_card("Analistas", f"{etit_n_analistas}", COR_INFO), unsafe_allow_html=True)
            with ek5:
                st.markdown(kpi_card("TMA Médio", f"{etit_tma_geral:.4f}", COR_PRIMARIA), unsafe_allow_html=True)
            with ek6:
                st.markdown(kpi_card("TMR Médio", f"{etit_tmr_geral:.4f}", COR_ALERTA), unsafe_allow_html=True)

            st.markdown("")

            # ---- Ranking por Analista ----
            st.markdown("##### 🏆 Ranking ETIT por Analista")
            resumo_etit = etit_resumo_analista(df_etit_filtrado)
            if not resumo_etit.empty:
                disp_etit = resumo_etit.copy()
                disp_etit["Nome"] = disp_etit["Nome"].apply(primeiro_nome)
                disp_cols_etit = ["Nome", "Setor", "Total_Eventos", "Eventos_Aderentes",
                                  "Aderencia_Pct", "RAL_Count", "REC_Count", "TMA_Medio", "TMR_Medio"]
                disp_cols_etit = [c for c in disp_cols_etit if c in disp_etit.columns]
                tbl_etit = disp_etit[disp_cols_etit].copy()
                tbl_etit.columns = [
                    c.replace("Total_Eventos", "Eventos")
                     .replace("Eventos_Aderentes", "Aderentes")
                     .replace("Aderencia_Pct", "Aderência %")
                     .replace("RAL_Count", "RAL")
                     .replace("REC_Count", "REC")
                     .replace("TMA_Medio", "TMA")
                     .replace("TMR_Medio", "TMR")
                    for c in disp_cols_etit
                ]
                tbl_etit = tbl_etit.reset_index(drop=True)
                tbl_etit.index += 1
                tbl_etit.index.name = "#"

                styled_etit = tbl_etit.style.format(
                    {"Aderência %": "{:.1f}", "TMA": "{:.4f}", "TMR": "{:.4f}"}, na_rep="—"
                )
                styled_etit = styled_etit.background_gradient(cmap="Purples", subset=["Eventos"])
                if "Aderência %" in tbl_etit.columns and tbl_etit["Aderência %"].notna().any():
                    styled_etit = styled_etit.background_gradient(
                        cmap="RdYlGn", subset=["Aderência %"], vmin=50, vmax=100
                    )
                st.dataframe(styled_etit, use_container_width=True)

                # Best/Worst cards
                if len(tbl_etit) >= 2:
                    best_etit = tbl_etit.iloc[0]
                    worst_etit = tbl_etit.iloc[-1]
                    best_ader_row = tbl_etit.sort_values("Aderência %", ascending=False).iloc[0] if "Aderência %" in tbl_etit.columns else None

                    ce1, ce2, ce3 = st.columns(3)
                    with ce1:
                        st.markdown(f"""<div class="perf-card perf-best">
                            <div class="p-title">🏆 Mais Eventos</div>
                            <div class="p-name" style="color:#2ECC71;">{best_etit['Nome']}</div>
                            <div class="p-detail">{best_etit['Eventos']:,.0f} eventos · Aderência: {best_etit.get('Aderência %', 0):.1f}%</div>
                        </div>""", unsafe_allow_html=True)
                    with ce2:
                        st.markdown(f"""<div class="perf-card perf-worst">
                            <div class="p-title">⚠️ Menos Eventos</div>
                            <div class="p-name" style="color:#E74C3C;">{worst_etit['Nome']}</div>
                            <div class="p-detail">{worst_etit['Eventos']:,.0f} eventos · Aderência: {worst_etit.get('Aderência %', 0):.1f}%</div>
                        </div>""", unsafe_allow_html=True)
                    with ce3:
                        if best_ader_row is not None:
                            st.markdown(f"""<div class="perf-card perf-dpa">
                                <div class="p-title">📊 Melhor Aderência</div>
                                <div class="p-name" style="color:#5DADE2;">{best_ader_row['Nome']}</div>
                                <div class="p-detail">Aderência: {best_ader_row['Aderência %']:.1f}% · {best_ader_row['Eventos']:,.0f} eventos</div>
                            </div>""", unsafe_allow_html=True)

            # Bar chart — eventos por analista
            st.markdown("")
            if not resumo_etit.empty:
                chart_etit = resumo_etit[["Nome", "Total_Eventos"]].copy()
                chart_etit["Nome"] = chart_etit["Nome"].apply(primeiro_nome)
                chart_etit = chart_etit.set_index("Nome").sort_values("Total_Eventos")
                chart_etit.columns = ["Eventos"]
                st.bar_chart(chart_etit, horizontal=True, color="#8E44AD", height=400)

            # ---- Breakdowns ----
            st.markdown("---")
            st.markdown("##### 📊 Análises de Composição ETIT")

            col_dem, col_tipo = st.columns(2)

            with col_dem:
                st.markdown("**Por Demanda (RAL/REC)**")
                dem = etit_por_demanda(df_etit_filtrado)
                if not dem.empty:
                    dem["Aderência %"] = (dem["Aderentes"] / dem["Eventos"] * 100).round(1)
                    dem["TMA_Medio"] = dem["TMA_Medio"].round(4)
                    dem["TMR_Medio"] = dem["TMR_Medio"].round(4)
                    st.dataframe(
                        dem.rename(columns={"TMA_Medio": "TMA", "TMR_Medio": "TMR"})
                           .style.format({"Aderência %": "{:.1f}", "TMA": "{:.4f}", "TMR": "{:.4f}"}, na_rep="—"),
                        use_container_width=True, hide_index=True,
                    )

            with col_tipo:
                st.markdown("**Por Tipo de Rede/Equipamento**")
                tipo = etit_por_tipo(df_etit_filtrado)
                if not tipo.empty:
                    tipo["Aderência %"] = (tipo["Aderentes"] / tipo["Eventos"] * 100).round(1)
                    st.dataframe(
                        tipo.style
                            .format({"Aderência %": "{:.1f}"}, na_rep="—")
                            .background_gradient(cmap="Purples", subset=["Eventos"]),
                        use_container_width=True, hide_index=True,
                    )

            col_causa, col_reg = st.columns(2)

            with col_causa:
                st.markdown("**Por Causa**")
                causa = etit_por_causa(df_etit_filtrado)
                if not causa.empty:
                    causa["Aderência %"] = (causa["Aderentes"] / causa["Eventos"] * 100).round(1)
                    st.dataframe(
                        causa.head(15).style
                            .format({"Aderência %": "{:.1f}"}, na_rep="—")
                            .background_gradient(cmap="Purples", subset=["Eventos"]),
                        use_container_width=True, hide_index=True,
                    )

            with col_reg:
                st.markdown("**Por Regional**")
                reg = etit_por_regional(df_etit_filtrado)
                if not reg.empty:
                    reg["Aderência %"] = (reg["Aderentes"] / reg["Eventos"] * 100).round(1)
                    st.dataframe(
                        reg.style
                            .format({"Aderência %": "{:.1f}"}, na_rep="—")
                            .background_gradient(cmap="Purples", subset=["Eventos"]),
                        use_container_width=True, hide_index=True,
                    )

            # Por Turno
            st.markdown("**Por Turno**")
            turno = etit_por_turno(df_etit_filtrado)
            if not turno.empty:
                turno["Aderência %"] = (turno["Aderentes"] / turno["Eventos"] * 100).round(1)
                col_t1, col_t2 = st.columns([1, 2])
                with col_t1:
                    st.dataframe(
                        turno.style.format({"Aderência %": "{:.1f}"}, na_rep="—"),
                        use_container_width=True, hide_index=True,
                    )
                with col_t2:
                    chart_turno = turno[["Turno", "Eventos"]].set_index("Turno")
                    st.bar_chart(chart_turno, color="#8E44AD", height=250)

            # ---- Evolução Diária ETIT ----
            st.markdown("---")
            st.markdown("##### 📅 Evolução Diária ETIT")
            daily_etit = etit_evolucao_diaria(df_etit_filtrado)
            if not daily_etit.empty:
                st.area_chart(daily_etit[["Data", "Eventos"]].set_index("Data"), color="#8E44AD", height=300)
                st.line_chart(daily_etit[["Data", "Aderencia_Pct"]].set_index("Data"), color=COR_SUCESSO, height=250)

            # ---- Dados Brutos ETIT ----
            st.markdown("---")
            st.markdown("##### 📋 Dados Brutos ETIT")
            etit_show_cols = [
                ETIT_COL_LOGIN, "Nome", "Setor", ETIT_COL_DEMANDA, ETIT_COL_NOTA,
                ETIT_COL_STATUS, ETIT_COL_TIPO, ETIT_COL_CAUSA,
                ETIT_COL_REGIONAL, ETIT_COL_CIDADE, ETIT_COL_UF,
                ETIT_COL_TURNO, ETIT_COL_TMA, ETIT_COL_TMR,
                ETIT_COL_DT_ACIONAMENTO, ETIT_COL_ANOMES,
            ]
            etit_show_cols = [c for c in etit_show_cols if c in df_etit_filtrado.columns]
            st.dataframe(
                df_etit_filtrado[etit_show_cols].sort_values(
                    [ETIT_COL_DT_ACIONAMENTO] if ETIT_COL_DT_ACIONAMENTO in df_etit_filtrado.columns else ["Nome"],
                    ascending=False,
                ),
                use_container_width=True, height=500,
            )

            csv_etit = df_etit_filtrado[etit_show_cols].to_csv(index=False).encode("utf-8")
            st.download_button(
                "📥 Baixar ETIT filtrado (CSV)",
                csv_etit, "etit_por_evento_equipe.csv", "text/csv",
            )


# =====================================================
# FOOTER
# =====================================================
st.markdown("---")
footer_parts = [
    f"Dashboard de Produtividade COP Rede",
    f"{len(df_filtrado)} registros",
    f"{n_analistas} analistas",
    f"Dados de {data_min.strftime('%d/%m/%Y') if pd.notna(data_min) else '?'} "
    f"a {data_max.strftime('%d/%m/%Y') if pd.notna(data_max) else '?'}",
]
if etit_loaded:
    footer_parts.append(f"ETIT: {len(df_etit_filtrado)} eventos")
st.caption(" · ".join(footer_parts))
