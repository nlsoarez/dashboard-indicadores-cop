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
    ETIT_COL_REGIONAL, ETIT_COL_GRUPO, ETIT_COL_TURNO, ETIT_COL_TMA, ETIT_COL_TMR,
    ETIT_COL_DT_ACIONAMENTO, ETIT_COL_ANOMES, ETIT_COL_INDICADOR_VAL,
    ETIT_COL_NOTA, ETIT_COL_AREA, ETIT_COL_CIDADE, ETIT_COL_UF,
    # Residencial Indicadores
    RES_INDICADORES_FILTRO, RES_IND_LABELS, RES_IND_COLORS,
    RES_IND_INVERTIDOS, RES_IND_ETIT_FIBRA_HFC, RES_IND_ETIT_GPON,
    RES_IND_REPROG_GPON, RES_IND_ASSERT_FIBRA_HFC, RES_IND_ASSERT_GPON,
    RES_COL_INDICADOR_NOME, RES_COL_VOLUME, RES_COL_INDICADOR_VAL as RES_COL_IND_VAL,
    RES_COL_STATUS, RES_COL_REGIONAL as RES_REGIONAL,
    RES_COL_DT_INICIO, RES_COL_TMA as RES_TMA, RES_COL_TMR as RES_TMR,
    RES_COL_SOLUCAO, RES_COL_IMPACTO, RES_COL_NATUREZA,
    RES_COL_GRUPO as RES_GRUPO,
    RES_COL_CIDADE, RES_COL_UF as RES_UF, RES_COL_ANOMES as RES_ANOMES,
    RES_COL_ID_MOSTRA,
    # DPA Ocupação
    DPA_THRESHOLD_OK, DPA_THRESHOLD_ALERTA,
    # Indicadores TOA
    TOA_IND_CANCELADAS, TOA_IND_VALIDACAO, TOA_IND_LABELS, TOA_IND_COLORS,
    TOA_INDICADORES_FILTRO, TOA_AGING_ORDER,
)
from src.processors import (
    load_produtividade, resumo_mensal, resumo_geral,
    evolucao_diaria, composicao_volume, primeiro_nome,
    # ETIT
    load_etit, etit_resumo_analista, etit_por_demanda,
    etit_por_tipo, etit_por_causa, etit_por_regional,
    etit_por_turno, etit_evolucao_diaria,
    # Residencial Indicadores
    load_residencial_indicadores,
    res_kpis_por_indicador, res_por_regional,
    res_por_natureza, res_por_solucao, res_por_impacto,
    res_evolucao_diaria,
    # DPA Ocupação
    load_dpa_ocupacao, dpa_ranking,
    # Indicadores TOA
    load_toa_indicadores, toa_resumo_por_indicador,
    toa_canceladas_por_analista, toa_canceladas_por_tipo,
    toa_canceladas_por_aging, toa_canceladas_por_rede, toa_canceladas_por_regional,
    toa_canceladas_evolucao,
    toa_validacao_por_analista, toa_validacao_por_tipo,
    toa_validacao_por_rede, toa_validacao_por_regional,
    toa_validacao_evolucao,
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
    .main-header {
        background: linear-gradient(135deg, #1B4F72 0%, #2980B9 100%);
        padding: 1.5rem 2rem; border-radius: 12px; margin-bottom: 1.5rem;
    }
    .main-header h1 { color: #fff !important; margin: 0; font-size: 1.8rem; }
    .main-header p  { color: #D6EAF8 !important; margin: 0.3rem 0 0 0; font-size: 0.95rem; }

    .kpi-card {
        background: var(--secondary-background-color); border-radius: 10px;
        padding: 1.2rem; border-left: 4px solid; text-align: center;
    }
    .kpi-card .kpi-value { font-size: 1.8rem; font-weight: 700; margin: 0.3rem 0; }
    .kpi-card .kpi-label { font-size: 0.78rem; opacity: 0.55; text-transform: uppercase; letter-spacing: 0.5px; }
    .kpi-card .kpi-delta { font-size: 0.8rem; margin-top: 0.2rem; }

    .section-header {
        font-size: 1.15rem; font-weight: 600; color: #5DADE2;
        border-bottom: 2px solid rgba(41,128,185,0.4);
        padding-bottom: 0.4rem; margin: 1.5rem 0 1rem 0;
    }

    .perf-card { padding: 0.9rem 1rem; border-radius: 10px; border-left: 4px solid; margin-bottom: 0.5rem; }
    .perf-best  { background: rgba(39,174,96,0.12);  border-left-color: #2ECC71; }
    .perf-worst { background: rgba(231,76,60,0.12);   border-left-color: #E74C3C; }
    .perf-dpa   { background: rgba(41,128,185,0.12);  border-left-color: #5DADE2; }
    .perf-card .p-title  { font-size: 0.8rem; font-weight: 600; margin-bottom: 0.25rem; }
    .perf-card .p-name   { font-size: 1.1rem; font-weight: 700; }
    .perf-card .p-detail { font-size: 0.82rem; opacity: 0.7; margin-top: 0.15rem; }

    .insight-card {
        background: var(--secondary-background-color); border-radius: 10px;
        padding: 0.85rem 1rem; margin-bottom: 0.6rem; border-left: 4px solid #5DADE2;
    }
    .tag-green {
        background: rgba(46,204,113,0.18); color: #2ECC71;
        padding: 2px 8px; border-radius: 10px;
        font-size: 0.73rem; font-weight: 500; display: inline-block; margin: 1px 2px;
    }
    .tag-red {
        background: rgba(231,76,60,0.18); color: #E74C3C;
        padding: 2px 8px; border-radius: 10px;
        font-size: 0.73rem; font-weight: 500; display: inline-block; margin: 1px 2px;
    }
    .sector-badge {
        background: rgba(93,173,226,0.15); color: #5DADE2;
        padding: 2px 8px; border-radius: 8px;
        font-size: 0.72rem; font-weight: 500; margin-left: 6px;
    }
    .rank-pill {
        background: rgba(255,255,255,0.06); padding: 2px 8px; border-radius: 8px;
        font-size: 0.72rem; opacity: 0.55; margin-left: 3px;
    }

    .leader-card {
        background: var(--secondary-background-color); border-radius: 12px;
        padding: 1rem 1.2rem; border-top: 3px solid #F1C40F; margin-bottom: 0.8rem;
    }
    .leader-card .l-name  { font-size: 1.05rem; font-weight: 700; }
    .leader-card .l-badge {
        background: rgba(241,196,15,0.15); color: #F1C40F;
        padding: 2px 8px; border-radius: 8px;
        font-size: 0.72rem; font-weight: 600; margin-left: 6px;
    }
    .leader-card .l-stat  { font-size: 0.82rem; opacity: 0.7; margin-top: 0.3rem; }
    .leader-card .l-vol   { font-size: 1.4rem; font-weight: 700; color: #5DADE2; }

    .sector-header {
        display: inline-block; padding: 0.35rem 1rem; border-radius: 8px;
        font-weight: 600; font-size: 0.9rem; margin-bottom: 0.6rem;
    }
    .sector-res { background: rgba(41,128,185,0.15); color: #5DADE2; }
    .sector-emp { background: rgba(243,156,18,0.15); color: #F39C12; }

    .etit-card {
        background: var(--secondary-background-color); border-radius: 10px;
        padding: 1rem 1.2rem; border-left: 4px solid #8E44AD; margin-bottom: 0.6rem;
    }

    .res-ind-card {
        background: var(--secondary-background-color); border-radius: 12px;
        padding: 1.1rem 1.2rem; border-top: 3px solid; margin-bottom: 0.6rem;
    }
    .res-ind-card .ri-title  { font-size: 0.8rem; font-weight: 600; opacity: 0.6; margin-bottom: 0.4rem; text-transform: uppercase; letter-spacing: 0.5px; }
    .res-ind-card .ri-vol    { font-size: 1.6rem; font-weight: 700; }
    .res-ind-card .ri-pct    { font-size: 1.1rem; font-weight: 600; margin-top: 0.15rem; }
    .res-ind-card .ri-detail { font-size: 0.78rem; opacity: 0.6; margin-top: 0.3rem; }

    /* DPA Ocupação cards */
    .dpa-card {
        background: var(--secondary-background-color); border-radius: 12px;
        padding: 1rem 1.2rem; border-left: 4px solid #16A085; margin-bottom: 0.5rem;
    }
    .dpa-card .dpa-nome  { font-size: 1rem; font-weight: 700; }
    .dpa-card .dpa-val   { font-size: 1.4rem; font-weight: 700; }
    .dpa-card .dpa-setor { font-size: 0.75rem; opacity: 0.6; margin-top: 0.15rem; }
    .dpa-semaforo-verde   { color: #27AE60; }
    .dpa-semaforo-amarelo { color: #F39C12; }
    .dpa-semaforo-vermelho{ color: #E74C3C; }

    .dataframe { font-size: 0.85rem !important; }

    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #1B4F72 0%, #154360 100%);
    }
    [data-testid="stSidebar"] * { color: white !important; }
    [data-testid="stSidebar"] .stSelectbox label { color: #D6EAF8 !important; }
</style>
""", unsafe_allow_html=True)


# (loaders são chamados diretamente — sem cache, para máxima confiabilidade)


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


def _dpa_color(pct):
    if pct is None or (isinstance(pct, float) and np.isnan(pct)):
        return COR_INFO
    if pct >= DPA_THRESHOLD_OK:
        return COR_SUCESSO
    if pct >= DPA_THRESHOLD_ALERTA:
        return COR_ALERTA
    return COR_PERIGO


def _dpa_semaforo(pct):
    if pct is None:
        return "—"
    if pct >= DPA_THRESHOLD_OK:
        return "🟢"
    if pct >= DPA_THRESHOLD_ALERTA:
        return "🟡"
    return "🔴"


def get_sector_vol_cols(setor, available_cols):
    cols = {}
    if setor in ("Todos", "RESIDENCIAL"):
        cols.update(VOL_COLS_RESIDENCIAL)
    if setor in ("Todos", "EMPRESARIAL"):
        cols.update(VOL_COLS_EMPRESARIAL)
    cols.update(VOL_COLS_AMBOS)
    return {k: v for k, v in cols.items() if k in available_cols}


def build_insights(resumo_df, setor_filter):
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
            "vol_total": row[COL_VOL_TOTAL], "media_diaria": row.get("Media_Diaria", 0),
            "dias": row.get("Dias_Trabalhados", 0), "vol_diff": vol_diff,
            "vol_rank": vol_rank, "dpa": dpa_val,
            "strengths": strengths[:4], "weaknesses": weaknesses[:4], "n_peers": n_peers,
        })
    return data


def render_insight_cards(insights_list):
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
        "Dias_Trabalhados": "Dias", "Media_Diaria": "Média/Dia", "DPA_Media": "DPA %",
    }
    rename.update({k: all_vol[k] for k in vol_keys})
    disp = disp.rename(columns=rename)
    disp = disp.sort_values("Vol. Total", ascending=False).reset_index(drop=True)
    disp.index += 1; disp.index.name = "#"
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
    if len(disp) >= 2:
        best = disp.iloc[0]; worst = disp.iloc[-1]; best_dpa_row = None
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
# UPLOAD
# =====================================================
with st.container():
    col_upload1, col_upload2, col_upload3, col_upload4, col_upload5, col_info = st.columns([2, 2, 2, 2, 2, 1])
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
    with col_upload3:
        uploaded_res_ind = st.file_uploader(
            "📁 Analítico Indicadores Residencial",
            type=["xlsx", "xls"],
            help="Planilha com indicadores ETIT Fibra HFC, ETIT GPON, Reprogramação GPON, Assertividade — opcional",
            key="upload_res_ind",
        )
    with col_upload4:
        uploaded_toa = st.file_uploader(
            "📁 Indicadores TOA",
            type=["xlsx", "xls"],
            help="Planilha Analitico_Indicadores_TOA com Tarefas Canceladas e Tempo de Validação — opcional",
            key="upload_toa",
        )
    with col_upload5:
        uploaded_dpa = st.file_uploader(
            "📁 Ocupação DPA 2026",
            type=["xlsx", "xls"],
            help=(
                "Planilha Ocupação_DPA_2026 com abas 'Consolidado' e 'Analistas'.\n"
                "Extrai automaticamente o mês mais recente com dados disponíveis. — opcional"
            ),
            key="upload_dpa",
        )
    with col_info:
        st.info(
            f"**Equipe monitorada:** {len(EQUIPE_IDS)} analistas\n\n"
            f"Empresarial: {len(BASE_EQUIPE[BASE_EQUIPE['Setor']=='EMPRESARIAL'])} · "
            f"Residencial: {len(BASE_EQUIPE[BASE_EQUIPE['Setor']=='RESIDENCIAL'])}"
        )

# Persistência no session_state
for key_name, file_obj in [
    ("uploaded_bytes",        uploaded_file),
    ("uploaded_etit_bytes",   uploaded_etit),
    ("uploaded_res_ind_bytes",uploaded_res_ind),
    ("uploaded_toa_bytes",    uploaded_toa),
    ("uploaded_dpa_bytes",    uploaded_dpa),
]:
    if file_obj is not None:
        st.session_state[key_name]          = file_obj.getvalue()
        st.session_state[key_name + "_name"] = file_obj.name

if "uploaded_bytes" not in st.session_state:
    st.markdown("---")
    st.markdown("### 👋 Bem-vindo!")
    st.markdown(
        "Faça upload da planilha **Produtividade COP Rede - Analítico** acima para "
        "visualizar os dados de produtividade da sua equipe.\n\n"
        "Opcionalmente, faça upload das planilhas adicionais:\n"
        "- **Analítico Empresarial** → dados de ETIT POR EVENTO\n"
        "- **Analítico Indicadores Residencial** → ETIT Fibra HFC, GPON, Reprog., Assertividade\n"
        "- **Ocupação DPA 2026** → DPA oficial por analista (mês mais recente detectado automaticamente)"
    )
    with st.expander("📋 Analistas monitorados"):
        st.dataframe(BASE_EQUIPE, use_container_width=True, hide_index=True)
    st.stop()


# =====================================================
# PROCESSAR DADOS — Produtividade
# =====================================================
try:
    with st.spinner("Carregando e processando dados de produtividade..."):
        df = load_produtividade(io.BytesIO(st.session_state["uploaded_bytes"]))
    if df.empty:
        st.error("Nenhum analista da equipe encontrado na planilha de produtividade.")
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
            df_etit = load_etit(io.BytesIO(st.session_state["uploaded_etit_bytes"]))
            etit_loaded = not df_etit.empty
            if not etit_loaded:
                st.warning("⚠️ ETIT: Planilha carregada mas nenhum analista da equipe encontrado.")
            else:
                st.toast(f"✅ ETIT carregado: {len(df_etit)} registros")
    except Exception as e:
        st.error(f"❌ Erro ao processar planilha ETIT: {e}")
        with st.expander("Detalhes do erro ETIT", expanded=True):
            st.code(traceback.format_exc())


# =====================================================
# PROCESSAR DADOS — Indicadores Residencial (opcional)
# =====================================================
df_res_ind = pd.DataFrame()
res_ind_loaded = False

if "uploaded_res_ind_bytes" in st.session_state:
    try:
        with st.spinner("Carregando Indicadores Residencial..."):
            df_res_ind = load_residencial_indicadores(io.BytesIO(st.session_state["uploaded_res_ind_bytes"]))
            res_ind_loaded = not df_res_ind.empty
            if not res_ind_loaded:
                st.warning("⚠️ Residencial: Planilha carregada mas nenhum indicador encontrado.")
            else:
                st.toast(f"✅ Residencial carregado: {len(df_res_ind)} registros")
    except Exception as e:
        st.error(f"❌ Erro ao processar planilha de Indicadores Residencial: {e}")
        with st.expander("Detalhes do erro Residencial", expanded=True):
            st.code(traceback.format_exc())


# =====================================================
# PROCESSAR DADOS — Ocupação DPA (opcional)
# =====================================================
df_dpa = pd.DataFrame()
dpa_mes_info = {}
dpa_loaded = False

if "uploaded_dpa_bytes" in st.session_state:
    try:
        with st.spinner("Carregando Ocupação DPA..."):
            df_dpa, dpa_mes_info = load_dpa_ocupacao(io.BytesIO(st.session_state["uploaded_dpa_bytes"]))
            dpa_loaded = not df_dpa.empty
            if not dpa_loaded:
                st.warning("⚠️ DPA: Planilha carregada mas nenhum analista da equipe encontrado.")
            else:
                st.toast(f"✅ DPA carregado: {len(df_dpa)} analistas")
    except Exception as e:
        st.error(f"❌ Erro ao processar planilha de Ocupação DPA: {e}")
        with st.expander("Detalhes do erro DPA", expanded=True):
            st.code(traceback.format_exc())


# =====================================================
# PROCESSAR DADOS — Indicadores TOA (opcional)
# =====================================================
df_toa = pd.DataFrame()
toa_loaded = False
toa_anomes = None

if "uploaded_toa_bytes" in st.session_state:
    try:
        with st.spinner("Carregando Indicadores TOA..."):
            df_toa = load_toa_indicadores(io.BytesIO(st.session_state["uploaded_toa_bytes"]))
            toa_loaded = not df_toa.empty
            if toa_loaded and "ANOMES" in df_toa.columns:
                toa_anomes = int(df_toa["ANOMES"].max())
            if not toa_loaded:
                st.warning("⚠️ TOA: Planilha carregada mas nenhum analista da equipe encontrado.")
            else:
                st.toast(f"✅ TOA carregado: {len(df_toa)} registros")
    except Exception as e:
        st.error(f"❌ Erro ao processar planilha de Indicadores TOA: {e}")
        with st.expander("Detalhes do erro TOA", expanded=True):
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
        for key in [
            "uploaded_bytes", "uploaded_bytes_name",
            "uploaded_etit_bytes", "uploaded_etit_bytes_name",
            "uploaded_res_ind_bytes", "uploaded_res_ind_bytes_name",
            "uploaded_toa_bytes", "uploaded_toa_bytes_name",
            "uploaded_dpa_bytes", "uploaded_dpa_bytes_name",
        ]:
            st.session_state.pop(key, None)
        st.rerun()

    st.markdown("---")
    st.markdown("### 📊 Equipe")
    analistas_options = df[[COL_LOGIN, COL_NOME]].drop_duplicates().sort_values(COL_NOME)
    analista_selecionado = st.selectbox(
        "Detalhe individual",
        options=["Todos"] + analistas_options[COL_LOGIN].tolist(),
        format_func=lambda x: "Visão Geral" if x == "Todos" else
            analistas_options[analistas_options[COL_LOGIN]==x][COL_NOME].iloc[0]
            if len(analistas_options[analistas_options[COL_LOGIN]==x]) > 0 else x,
    )

    st.markdown("---")
    st.markdown("### 📋 Status dos dados")

    # Diagnóstico de uploads
    _upload_keys = {
        "ETIT": "uploaded_etit_bytes",
        "Residencial": "uploaded_res_ind_bytes",
        "TOA": "uploaded_toa_bytes",
        "DPA": "uploaded_dpa_bytes",
    }
    for _label, _key in _upload_keys.items():
        if _key in st.session_state:
            _sz = len(st.session_state[_key])
            st.caption(f"📎 {_label}: {_sz:,} bytes carregados")
        else:
            st.caption(f"⬜ {_label}: não carregado")

    if etit_loaded:
        st.success(f"✅ ETIT: {len(df_etit)} eventos")
    elif "uploaded_etit_bytes" in st.session_state:
        st.warning("⚠️ ETIT: planilha carregada mas sem dados da equipe")
    if res_ind_loaded:
        st.success(f"✅ Ind. Residencial: {len(df_res_ind):,} registros")
    elif "uploaded_res_ind_bytes" in st.session_state:
        st.warning("⚠️ Residencial: planilha carregada mas sem dados")
    if toa_loaded:
        anomes_str = str(toa_anomes) if toa_anomes else "?"
        n_canc = len(df_toa[df_toa["INDICADOR_NOME"] == "TAREFAS CANCELADAS"]) if toa_loaded else 0
        n_val  = len(df_toa[df_toa["INDICADOR_NOME"] == "TEMPO DE VALIDAÇÃO DO FORMULÁRIO"]) if toa_loaded else 0
        st.success(f"✅ TOA {anomes_str}: {n_canc} canceladas · {n_val} validações")
    elif "uploaded_toa_bytes" in st.session_state:
        st.warning("⚠️ TOA: planilha carregada mas sem dados da equipe")
    if dpa_loaded:
        mes_label = dpa_mes_info.get("mes_nome", "?")
        dpa_geral = dpa_mes_info.get("dpa_geral_pct")
        dpa_g_str = f" · {dpa_geral:.1f}%" if dpa_geral else ""
        st.success(f"✅ Ocup. DPA: {len(df_dpa)} analistas · {mes_label}{dpa_g_str}")
    elif "uploaded_dpa_bytes" in st.session_state:
        st.warning("⚠️ DPA: planilha carregada mas sem dados da equipe")

    # Filtro Indicadores Residencial
    if res_ind_loaded:
        st.markdown("### 🏠 Filtro Indicadores Res.")
        res_ind_selecionado = st.selectbox(
            "Indicador",
            options=["Todos"] + RES_INDICADORES_FILTRO,
            format_func=lambda x: "Todos os indicadores" if x == "Todos" else RES_IND_LABELS.get(x, x),
            key="res_ind_filter",
        )
        if RES_ANOMES in df_res_ind.columns:
            res_meses = sorted(df_res_ind[RES_ANOMES].dropna().unique().tolist())
            res_mes_sel = st.selectbox(
                "Período (Indicadores)",
                options=["Todos"] + res_meses,
                format_func=lambda x: "Todos" if x == "Todos" else str(x),
                key="res_mes_filter",
            )
        else:
            res_mes_sel = "Todos"
    else:
        res_ind_selecionado = "Todos"
        res_mes_sel = "Todos"


# =====================================================
# APLICAR FILTROS
# =====================================================
df_filtrado = df.copy()
if mes_selecionado != "Todos":
    df_filtrado = df_filtrado[df_filtrado[COL_ANOMES] == mes_selecionado]
if setor_selecionado != "Todos":
    df_filtrado = df_filtrado[df_filtrado["Setor"] == setor_selecionado]

df_etit_filtrado = df_etit.copy()
if etit_loaded:
    if mes_selecionado != "Todos" and ETIT_COL_ANOMES in df_etit_filtrado.columns:
        df_etit_filtrado = df_etit_filtrado[df_etit_filtrado[ETIT_COL_ANOMES] == str(mes_selecionado)]
    if setor_selecionado != "Todos":
        df_etit_filtrado = df_etit_filtrado[df_etit_filtrado["Setor"] == setor_selecionado]
    if analista_selecionado != "Todos":
        df_etit_filtrado = df_etit_filtrado[df_etit_filtrado[ETIT_COL_LOGIN] == analista_selecionado]

df_res_filtrado = df_res_ind.copy()
if res_ind_loaded:
    if res_mes_sel != "Todos" and RES_ANOMES in df_res_filtrado.columns:
        df_res_filtrado = df_res_filtrado[df_res_filtrado[RES_ANOMES] == str(res_mes_sel)]
    if res_ind_selecionado != "Todos":
        df_res_filtrado = df_res_filtrado[df_res_filtrado[RES_COL_INDICADOR_NOME] == res_ind_selecionado]

# DPA não precisa de filtro — já é o mês mais recente detectado automaticamente
df_dpa_filtrado = df_dpa.copy()
if dpa_loaded and setor_selecionado != "Todos":
    df_dpa_filtrado = df_dpa_filtrado[df_dpa_filtrado["Setor"] == setor_selecionado]


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

# KPI de DPA Oficial (se carregado)
dpa_oficial_geral = dpa_mes_info.get("dpa_geral_pct") if dpa_loaded else None

n_kpi_cols = 6 if dpa_loaded else 5
kpi_cols = st.columns(n_kpi_cols)

with kpi_cols[0]:
    st.markdown(kpi_card("Volume Total", f"{total_vol:,.0f}", COR_PRIMARIA), unsafe_allow_html=True)
with kpi_cols[1]:
    st.markdown(kpi_card("Analistas Ativos", f"{n_analistas}", COR_INFO), unsafe_allow_html=True)
with kpi_cols[2]:
    st.markdown(kpi_card("Média/Analista", f"{media_diaria_equipe:,.0f}", COR_SUCESSO), unsafe_allow_html=True)
with kpi_cols[3]:
    dpa_display = f"{dpa_media:.0f}" if dpa_media is not None else "—"
    dpa_color = _dpa_color(dpa_media)
    st.markdown(kpi_card("DPA Calc. Média", dpa_display, dpa_color, suffix="%"), unsafe_allow_html=True)
with kpi_cols[4]:
    periodo_str = f"{data_min.strftime('%d/%m')}" if pd.notna(data_min) else "—"
    periodo_str += f" a {data_max.strftime('%d/%m')}" if pd.notna(data_max) else ""
    st.markdown(kpi_card("Período", periodo_str, COR_INFO), unsafe_allow_html=True)

if dpa_loaded:
    with kpi_cols[5]:
        mes_nome = dpa_mes_info.get("mes_nome", "?")
        dpa_of_display = f"{dpa_oficial_geral:.1f}" if dpa_oficial_geral else "—"
        dpa_of_color = _dpa_color(dpa_oficial_geral)
        st.markdown(
            kpi_card(f"DPA Oficial ({mes_nome[:3]})", dpa_of_display, dpa_of_color, suffix="%"),
            unsafe_allow_html=True,
        )


# =====================================================
# VISÃO INDIVIDUAL
# =====================================================
if analista_selecionado != "Todos":
    df_analista = df_filtrado[df_filtrado[COL_LOGIN] == analista_selecionado]
    if not df_analista.empty:
        nome_analista = df_analista[COL_NOME].iloc[0]
        st.markdown(f'<div class="section-header">👤 Detalhe: {nome_analista}</div>', unsafe_allow_html=True)

        ca1, ca2, ca3, ca4, ca5 = st.columns(5)
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
            st.markdown(kpi_card("DPA Calc.", dpa_d, COR_ALERTA, suffix="%"), unsafe_allow_html=True)
        with ca5:
            if dpa_loaded:
                dpa_row = df_dpa[df_dpa["Login"] == analista_selecionado]
                dpa_of_ind = dpa_row["DPA_Pct_Oficial"].iloc[0] if not dpa_row.empty else None
                dpa_of_str = f"{dpa_of_ind:.1f}" if dpa_of_ind else "—"
                mes_nome = dpa_mes_info.get("mes_nome", "")[:3]
                st.markdown(
                    kpi_card(f"DPA Oficial ({mes_nome})", dpa_of_str, _dpa_color(dpa_of_ind), suffix="%"),
                    unsafe_allow_html=True,
                )

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
if res_ind_loaded:
    tab_labels.append("🏠 Indicadores Residencial")
if toa_loaded:
    tab_labels.append("📋 Indicadores TOA")
if dpa_loaded:
    tab_labels.append("📊 Ocupação DPA")

tabs = st.tabs(tab_labels)

_base_tabs = 5
_tab_etit_idx  = _base_tabs if etit_loaded else None
_tab_res_idx   = (_base_tabs + (1 if etit_loaded else 0)) if res_ind_loaded else None
_tab_toa_idx   = (
    _base_tabs + (1 if etit_loaded else 0) + (1 if res_ind_loaded else 0)
) if toa_loaded else None
_tab_dpa_idx   = (
    _base_tabs + (1 if etit_loaded else 0) + (1 if res_ind_loaded else 0) + (1 if toa_loaded else 0)
) if dpa_loaded else None


# ---- TAB 1: RANKING ----
with tabs[0]:
    resumo = resumo_geral(df_filtrado)
    if not resumo.empty:
        resumo_equipe = resumo[~resumo[COL_LOGIN].isin(LIDERES_IDS)].copy()

        col_rank1, col_rank2 = st.columns(2)
        with col_rank1:
            st.markdown("#### 📦 Ranking por Volume Total")
            rank_vol = resumo_equipe[[COL_LOGIN, COL_NOME, "Setor", COL_VOL_TOTAL, "Dias_Trabalhados", "Media_Diaria"]].copy()
            rank_vol["Nome"] = rank_vol[COL_NOME].apply(primeiro_nome)
            rank_vol = rank_vol.sort_values(COL_VOL_TOTAL, ascending=False).reset_index(drop=True)
            rank_vol.index += 1; rank_vol.index.name = "#"
            display_vol = rank_vol[["Nome", "Setor", COL_VOL_TOTAL, "Dias_Trabalhados", "Media_Diaria"]].copy()
            display_vol.columns = ["Analista", "Setor", "Vol. Total", "Dias", "Média/Dia"]
            st.dataframe(
                display_vol.style.background_gradient(cmap="Blues", subset=["Vol. Total"]),
                use_container_width=True, height=500,
            )

        with col_rank2:
            st.markdown("#### ⏱️ Ranking por DPA")

            # Se DPA Oficial carregado, mostrar esse ranking; senão, o calculado
            if dpa_loaded and not df_dpa_filtrado.empty:
                rank_dpa_of = dpa_ranking(df_dpa_filtrado)
                rank_dpa_of = rank_dpa_of[~rank_dpa_of["Login"].isin(LIDERES_IDS)].reset_index(drop=True)
                rank_dpa_of.index += 1; rank_dpa_of.index.name = "#"
                _dpa_g = dpa_mes_info.get('dpa_geral_pct')
                _dpa_g_str = f"{_dpa_g:.1f}%" if _dpa_g is not None else "?"
                st.caption(
                    f"📌 DPA Oficial — {dpa_mes_info.get('mes_nome') or '?'} "
                    f"(média geral da equipe: {_dpa_g_str})"
                )

                # Adicionar semáforo
                rank_dpa_of["Status"] = rank_dpa_of["DPA %"].apply(_dpa_semaforo)
                rank_dpa_of = rank_dpa_of[["Status", "Analista", "Setor", "DPA %"]]
                st.dataframe(
                    rank_dpa_of.style
                        .format({"DPA %": "{:.1f}"})
                        .background_gradient(cmap="RdYlGn", subset=["DPA %"], vmin=50, vmax=100),
                    use_container_width=True, height=500,
                )
            else:
                rank_dpa = resumo_equipe[[COL_LOGIN, COL_NOME, "Setor", "DPA_Media"]].copy()
                rank_dpa["Nome"] = rank_dpa[COL_NOME].apply(primeiro_nome)
                rank_dpa = rank_dpa.dropna(subset=["DPA_Media"])
                rank_dpa = rank_dpa.sort_values("DPA_Media", ascending=False).reset_index(drop=True)
                rank_dpa.index += 1; rank_dpa.index.name = "#"
                display_dpa = rank_dpa[["Nome", "Setor", "DPA_Media"]].copy()
                display_dpa.columns = ["Analista", "Setor", "DPA %"]
                st.dataframe(
                    display_dpa.style
                        .format({"DPA %": "{:.1f}"})
                        .background_gradient(cmap="RdYlGn", subset=["DPA %"], vmin=50, vmax=100),
                    use_container_width=True, height=500,
                )

        st.markdown("#### 📊 Ranking por Média Diária")
        rank_media = resumo_equipe[[COL_NOME, "Media_Diaria"]].copy()
        rank_media["Nome"] = rank_media[COL_NOME].apply(primeiro_nome)
        chart_data = rank_media[["Nome", "Media_Diaria"]].set_index("Nome").sort_values("Media_Diaria")
        st.bar_chart(chart_data, horizontal=True, color=COR_PRIMARIA, height=500)

        st.markdown("---")
        st.markdown("#### 📋 Análise Detalhada por Setor")
        if setor_selecionado in ("Todos", "RESIDENCIAL"):
            render_sector_table(resumo_equipe, "RESIDENCIAL", VOL_COLS_RESIDENCIAL, "Blues")
        if setor_selecionado in ("Todos", "EMPRESARIAL"):
            render_sector_table(resumo_equipe, "EMPRESARIAL", VOL_COLS_EMPRESARIAL, "Oranges")

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

        cols_lid = st.columns(len(resumo_lid))
        for idx, (_, lrow) in enumerate(resumo_lid.sort_values(COL_VOL_TOTAL, ascending=False).iterrows()):
            nome_l = primeiro_nome(lrow[COL_NOME])
            setor_l = lrow["Setor"]
            vol_l = lrow[COL_VOL_TOTAL]
            media_l = lrow.get("Media_Diaria", 0)
            dias_l = lrow.get("Dias_Trabalhados", 0)
            dpa_l = lrow.get("DPA_Media", None)

            # DPA Oficial para o líder
            dpa_of_l = None
            if dpa_loaded:
                dpa_row = df_dpa[df_dpa["Login"] == lrow[COL_LOGIN]]
                if not dpa_row.empty:
                    dpa_of_l = dpa_row["DPA_Pct_Oficial"].iloc[0]

            dpa_l_str = f"{dpa_of_l:.1f}% (oficial)" if dpa_of_l else (f"{dpa_l:.1f}%" if pd.notna(dpa_l) else "—")

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
        det_lid.index += 1; det_lid.index.name = "#"
        fmt_lid = {"DPA %": "{:.1f}", "Média/Dia": "{:.1f}"}
        styled_lid = det_lid.style.format(fmt_lid, na_rep="—")
        styled_lid = styled_lid.background_gradient(cmap="YlOrBr", subset=["Vol. Total"])
        if "DPA %" in det_lid.columns and det_lid["DPA %"].notna().any():
            styled_lid = styled_lid.background_gradient(cmap="RdYlGn", subset=["DPA %"], vmin=50, vmax=100)
        st.dataframe(styled_lid, use_container_width=True)

        st.markdown("---")
        st.markdown("#### 💡 Insights dos Líderes")
        lid_insights = build_insights(resumo_full, setor_selecionado)
        lid_insights = [i for i in lid_insights if i["login"] in LIDERES_IDS]
        render_insight_cards(lid_insights)
    else:
        st.info("Nenhum líder encontrado nos dados filtrados.")

    # ---- ETIT dos Líderes ----
    if etit_loaded and not df_etit_filtrado.empty:
        _etit_lid = df_etit_filtrado[df_etit_filtrado[ETIT_COL_LOGIN].isin(LIDERES_IDS)].copy()
        if not _etit_lid.empty:
            st.markdown("---")
            st.markdown("#### ⚡ ETIT dos Líderes")
            _META_ETIT = 90.0
            for _dem_tipo in ["RAL", "REC"]:
                _dem_df = _etit_lid[_etit_lid[ETIT_COL_DEMANDA] == _dem_tipo]
                if _dem_df.empty:
                    continue
                st.markdown(f"##### {_dem_tipo} — Meta {_META_ETIT:.0f}%")
                _lid_etit_grp = _dem_df.groupby(ETIT_COL_LOGIN).agg(
                    Volume=(ETIT_COL_VOLUME, "sum") if ETIT_COL_VOLUME in _dem_df.columns else (ETIT_COL_STATUS, "count"),
                    Aderentes=(ETIT_COL_STATUS, lambda x: (x == "ADERENTE").sum()),
                ).reset_index()
                _lid_etit_grp["Aderência %"] = (_lid_etit_grp["Aderentes"] / _lid_etit_grp["Volume"] * 100).round(1)
                # Map login to name
                _login_nome = _dem_df.drop_duplicates(ETIT_COL_LOGIN)[[ETIT_COL_LOGIN, "Nome"]].set_index(ETIT_COL_LOGIN)["Nome"]
                _lid_etit_grp["Analista"] = _lid_etit_grp[ETIT_COL_LOGIN].map(_login_nome).apply(primeiro_nome)
                _lid_etit_grp = _lid_etit_grp.sort_values("Aderência %", ascending=False).reset_index(drop=True)
                _lid_etit_grp.index += 1; _lid_etit_grp.index.name = "#"
                _lid_etit_show = _lid_etit_grp[["Analista", "Volume", "Aderentes", "Aderência %"]]
                st.dataframe(
                    _lid_etit_show.style
                        .format({"Aderência %": "{:.1f}"}, na_rep="—")
                        .background_gradient(cmap="RdYlGn", subset=["Aderência %"], vmin=50, vmax=100),
                    use_container_width=True,
                )

    # ---- TOA dos Líderes ----
    if toa_loaded and not df_toa.empty:
        _toa_lid = df_toa[(df_toa["Setor"] == "EMPRESARIAL") & (df_toa["LOGIN"].isin(LIDERES_IDS))].copy()
        if not _toa_lid.empty:
            st.markdown("---")
            st.markdown("#### 📋 TOA dos Líderes")

            # Canceladas
            _toa_lid_canc = _toa_lid[_toa_lid["INDICADOR_NOME"] == TOA_IND_CANCELADAS]
            if not _toa_lid_canc.empty:
                st.markdown("##### ❌ Canceladas")
                _lc = _toa_lid_canc.groupby(["LOGIN", "Nome"]).size().reset_index(name="Canceladas")
                _lc["Analista"] = _lc["Nome"].apply(primeiro_nome)
                _lc = _lc.sort_values("Canceladas", ascending=True).reset_index(drop=True)
                _lc.index += 1; _lc.index.name = "#"
                st.dataframe(
                    _lc[["Analista", "Canceladas"]].style
                        .background_gradient(cmap="Reds", subset=["Canceladas"]),
                    use_container_width=True,
                )

            # Validação
            _toa_lid_val = _toa_lid[_toa_lid["INDICADOR_NOME"] == TOA_IND_VALIDACAO]
            if not _toa_lid_val.empty:
                st.markdown("##### ✅ Validação do Formulário")
                _lv = _toa_lid_val.groupby(["LOGIN", "Nome"]).agg(
                    Total=("INDICADOR", "count"),
                    Aderentes=("ADERENTE", "sum"),
                ).reset_index()
                _lv["Aderência %"] = (_lv["Aderentes"] / _lv["Total"] * 100).round(1)
                if "TMR_min" in _toa_lid_val.columns:
                    _lv_tmr = _toa_lid_val.groupby("LOGIN")["TMR_min"].mean().reset_index()
                    _lv_tmr.columns = ["LOGIN", "TMR (min)"]
                    _lv = _lv.merge(_lv_tmr, on="LOGIN", how="left")
                _lv["Analista"] = _lv["Nome"].apply(primeiro_nome)
                _lv = _lv.sort_values("Aderência %", ascending=False).reset_index(drop=True)
                _lv.index += 1; _lv.index.name = "#"
                _lv_cols = ["Analista", "Total", "Aderentes", "Aderência %"]
                _lv_fmt = {"Aderência %": "{:.1f}"}
                if "TMR (min)" in _lv.columns:
                    _lv_cols.append("TMR (min)")
                    _lv_fmt["TMR (min)"] = "{:.1f}"
                st.dataframe(
                    _lv[_lv_cols].style
                        .format(_lv_fmt, na_rep="—")
                        .background_gradient(cmap="RdYlGn", subset=["Aderência %"], vmin=40, vmax=100),
                    use_container_width=True,
                )

    # ---- DPA dos Líderes ----
    if dpa_loaded and not df_dpa_filtrado.empty:
        _dpa_lid = df_dpa_filtrado[df_dpa_filtrado["Login"].isin(LIDERES_IDS)].copy()
        if not _dpa_lid.empty:
            st.markdown("---")
            st.markdown("#### 📊 DPA dos Líderes")
            _dpa_lid["Analista"] = _dpa_lid["Nome"].apply(primeiro_nome)
            _dpa_lid = _dpa_lid.sort_values("DPA_Pct_Oficial", ascending=False).reset_index(drop=True)
            _dpa_lid.index += 1; _dpa_lid.index.name = "#"
            _dpa_lid_show = _dpa_lid[["Analista", "Setor", "DPA_Pct_Oficial"]].copy()
            _dpa_lid_show.columns = ["Analista", "Setor", "DPA %"]
            st.dataframe(
                _dpa_lid_show.style
                    .format({"DPA %": "{:.1f}"})
                    .background_gradient(cmap="RdYlGn", subset=["DPA %"], vmin=50, vmax=100),
                use_container_width=True,
            )


# ---- TAB 2: EVOLUÇÃO DIÁRIA ----
with tabs[2]:
    daily = evolucao_diaria(df_filtrado)
    if not daily.empty:
        st.markdown("#### Volume Total da Equipe por Dia")
        st.area_chart(daily[[COL_DATA, "Vol_Total"]].set_index(COL_DATA), color=COR_INFO, height=350)
        st.markdown("#### Média por Analista por Dia")
        st.line_chart(daily[[COL_DATA, "Media_por_Analista"]].set_index(COL_DATA), color=COR_SUCESSO, height=300)
        st.markdown("#### Analistas Ativos por Dia")
        st.bar_chart(daily[[COL_DATA, "Analistas"]].set_index(COL_DATA), color=COR_ALERTA, height=250)

    meses_unicos = df_filtrado[COL_MES].dropna().unique()
    if len(meses_unicos) > 1 and mes_selecionado == "Todos":
        st.markdown("---")
        st.markdown("#### 📅 Comparação Mensal por Analista")
        mensal = resumo_mensal(df_filtrado)
        if not mensal.empty:
            pivot = mensal.pivot_table(
                index=COL_NOME, columns=COL_MES, values="Media_Diaria", aggfunc="first"
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
                comp[["Atividade", "Volume", "Percentual"]].style.background_gradient(cmap="YlOrRd", subset=["Volume"]),
                use_container_width=True, hide_index=True,
            )
        with col_bar:
            st.bar_chart(comp[["Atividade", "Volume"]].set_index("Atividade").sort_values("Volume"),
                         horizontal=True, color=COR_PRIMARIA, height=450)

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
        display_labels = ["Analista", "Setor", "Dias", "Vol. Total", "Média/Dia", "DPA Calc. %"]
        sector_vol = get_sector_vol_cols(setor_selecionado, resumo_det.columns)
        vol_keys = list(sector_vol.keys()); vol_labels = list(sector_vol.values())
        available_base = [c for c in display_cols if c in resumo_det.columns]
        available_labels = display_labels[:len(available_base)]
        det = resumo_det[available_base + vol_keys].copy()
        det.columns = available_labels + vol_labels
        det = det.sort_values("Vol. Total", ascending=False).reset_index(drop=True)
        det.index += 1; det.index.name = "#"
        format_dict = {}
        if "DPA Calc. %" in det.columns:
            format_dict["DPA Calc. %"] = "{:.1f}"
        st.dataframe(det.style.format(format_dict) if format_dict else det, use_container_width=True)

    st.markdown("---")
    st.markdown("#### Dados Brutos (Filtrados)")
    cols_to_show = [COL_LOGIN, COL_NOME, "Setor", COL_DATA, COL_MES, COL_VOL_TOTAL, COL_DPA_RESULTADO]
    vol_cols_existing = [c for c in VOL_COLS.keys() if c in df_filtrado.columns]
    cols_to_show += vol_cols_existing
    cols_existing = [c for c in cols_to_show if c in df_filtrado.columns]
    st.dataframe(df_filtrado[cols_existing].sort_values([COL_NOME, COL_DATA]),
                 use_container_width=True, height=500)
    csv = df_filtrado[cols_existing].to_csv(index=False).encode("utf-8")
    st.download_button("📥 Baixar dados filtrados (CSV)", csv, "produtividade_equipe.csv", "text/csv")


# ---- TAB 5: ETIT POR EVENTO ----
if etit_loaded and _tab_etit_idx is not None:
    with tabs[_tab_etit_idx]:
        # Excluir líderes da análise principal
        _etit_eq = df_etit_filtrado[~df_etit_filtrado[ETIT_COL_LOGIN].isin(LIDERES_IDS)].copy()

        st.markdown("#### ⚡ ETIT POR EVENTO — Análise da Equipe")
        st.caption("Dados dos analistas (sem líderes). Meta de aderência: **90%** para RAL e REC.")

        if _etit_eq.empty:
            st.warning("Nenhum dado ETIT POR EVENTO encontrado com os filtros atuais.")
        else:
            # ---- Seção RAL e REC separadas ----
            _META_ETIT = 90.0
            for _dem_tipo in ["RAL", "REC"]:
                _dem_df = _etit_eq[_etit_eq[ETIT_COL_DEMANDA] == _dem_tipo]
                if _dem_df.empty:
                    continue
                _dem_vol = _dem_df[ETIT_COL_VOLUME].sum()
                _dem_ader = _dem_df[ETIT_COL_INDICADOR_VAL].sum()
                _dem_pct = (_dem_ader / _dem_vol * 100) if _dem_vol > 0 else 0
                _dem_tma = _dem_df[ETIT_COL_TMA].mean()
                _dem_tmr = _dem_df[ETIT_COL_TMR].mean()
                _dem_n = _dem_df[ETIT_COL_LOGIN].nunique()
                _meta_ok = _dem_pct >= _META_ETIT
                _meta_icon = "✅" if _meta_ok else "❌"
                _meta_color = COR_SUCESSO if _meta_ok else COR_PERIGO

                st.markdown(f"### {_meta_icon} {_dem_tipo} — Aderência: **{_dem_pct:.1f}%** (Meta: {_META_ETIT:.0f}%)")
                dk1, dk2, dk3, dk4, dk5 = st.columns(5)
                with dk1:
                    st.markdown(kpi_card(f"Eventos {_dem_tipo}", f"{_dem_vol:,.0f}", "#8E44AD"), unsafe_allow_html=True)
                with dk2:
                    st.markdown(kpi_card("Aderentes", f"{_dem_ader:,.0f}", COR_SUCESSO), unsafe_allow_html=True)
                with dk3:
                    st.markdown(kpi_card("Aderência", f"{_dem_pct:.1f}", _meta_color, suffix="%"), unsafe_allow_html=True)
                with dk4:
                    st.markdown(kpi_card("TMA Médio", f"{_dem_tma:.4f}", COR_PRIMARIA), unsafe_allow_html=True)
                with dk5:
                    st.markdown(kpi_card("Analistas", f"{_dem_n}", COR_INFO), unsafe_allow_html=True)

                # Ranking por analista para esta demanda
                _dem_rank = _dem_df.groupby([ETIT_COL_LOGIN, "Nome", "Setor"]).agg(
                    Eventos=(ETIT_COL_VOLUME, "sum"),
                    Aderentes=(ETIT_COL_INDICADOR_VAL, "sum"),
                    TMA=(ETIT_COL_TMA, "mean"),
                    TMR=(ETIT_COL_TMR, "mean"),
                ).reset_index()
                _dem_rank["Aderência %"] = (_dem_rank["Aderentes"] / _dem_rank["Eventos"] * 100).round(1)
                _dem_rank["Nome"] = _dem_rank["Nome"].apply(primeiro_nome)
                _dem_rank["TMA"] = _dem_rank["TMA"].round(4)
                _dem_rank["TMR"] = _dem_rank["TMR"].round(4)
                _dem_rank = _dem_rank.sort_values("Eventos", ascending=False).reset_index(drop=True)
                _dem_rank.index += 1; _dem_rank.index.name = "#"

                col_rk, col_gr = st.columns(2)
                with col_rk:
                    st.markdown(f"**Ranking {_dem_tipo} por Analista**")
                    _dem_show = _dem_rank[["Nome", "Setor", "Eventos", "Aderentes", "Aderência %", "TMA", "TMR"]]
                    _sty = _dem_show.style.format({"Aderência %": "{:.1f}", "TMA": "{:.4f}", "TMR": "{:.4f}"}, na_rep="—")
                    _sty = _sty.background_gradient(cmap="RdYlGn", subset=["Aderência %"], vmin=50, vmax=100)
                    _sty = _sty.background_gradient(cmap="Purples", subset=["Eventos"])
                    st.dataframe(_sty, use_container_width=True)

                with col_gr:
                    # IN_GRUPO para esta demanda
                    if ETIT_COL_GRUPO in _dem_df.columns:
                        st.markdown(f"**{_dem_tipo} por Grupo (IN_GRUPO)**")
                        _g = _dem_df.groupby(ETIT_COL_GRUPO).agg(
                            Eventos=(ETIT_COL_VOLUME, "sum"),
                            Aderentes=(ETIT_COL_INDICADOR_VAL, "sum"),
                        ).reset_index().rename(columns={ETIT_COL_GRUPO: "Grupo"})
                        _g["Aderência %"] = (_g["Aderentes"] / _g["Eventos"] * 100).round(1)
                        _g = _g.sort_values("Eventos", ascending=False).reset_index(drop=True)
                        if not _g.empty:
                            _best_g = _g.loc[_g["Aderência %"].idxmax()]
                            _worst_g = _g.loc[_g["Aderência %"].idxmin()]
                            st.caption(
                                f"🟢 Melhor: **{_best_g['Grupo']}** ({_best_g['Aderência %']:.1f}%) · "
                                f"🔴 Pior: **{_worst_g['Grupo']}** ({_worst_g['Aderência %']:.1f}%)"
                            )
                            _sg = _g.style.format({"Aderência %": "{:.1f}"}, na_rep="—")
                            _sg = _sg.background_gradient(cmap="RdYlGn", subset=["Aderência %"], vmin=50, vmax=100)
                            _sg = _sg.background_gradient(cmap="Purples", subset=["Eventos"])
                            st.dataframe(_sg, use_container_width=True, hide_index=True)

                st.markdown("---")

            # ---- IN_GRUPO geral ----
            if ETIT_COL_GRUPO in _etit_eq.columns:
                st.markdown("##### 📍 Visão Geral por Grupo (IN_GRUPO)")
                _gg = _etit_eq.groupby(ETIT_COL_GRUPO).agg(
                    Eventos=(ETIT_COL_VOLUME, "sum"),
                    Aderentes=(ETIT_COL_INDICADOR_VAL, "sum"),
                    TMA=(ETIT_COL_TMA, "mean"),
                    TMR=(ETIT_COL_TMR, "mean"),
                ).reset_index().rename(columns={ETIT_COL_GRUPO: "Grupo"})
                _gg["Aderência %"] = (_gg["Aderentes"] / _gg["Eventos"] * 100).round(1)
                _gg["TMA"] = _gg["TMA"].round(4)
                _gg["TMR"] = _gg["TMR"].round(4)
                _gg = _gg.sort_values("Eventos", ascending=False).reset_index(drop=True)
                if not _gg.empty:
                    _bg = _gg.loc[_gg["Aderência %"].idxmax()]
                    _wg = _gg.loc[_gg["Aderência %"].idxmin()]
                    st.caption(
                        f"🟢 Melhor grupo: **{_bg['Grupo']}** ({_bg['Aderência %']:.1f}%) · "
                        f"🔴 Pior grupo: **{_wg['Grupo']}** ({_wg['Aderência %']:.1f}%)"
                    )
                    _sgg = _gg.style.format({"Aderência %": "{:.1f}", "TMA": "{:.4f}", "TMR": "{:.4f}"}, na_rep="—")
                    _sgg = _sgg.background_gradient(cmap="RdYlGn", subset=["Aderência %"], vmin=50, vmax=100)
                    st.dataframe(_sgg, use_container_width=True, hide_index=True)
                st.markdown("---")

            # ---- Breakdowns existentes ----
            col_tipo, col_causa = st.columns(2)
            with col_tipo:
                st.markdown("**Por Tipo**")
                tipo = etit_por_tipo(_etit_eq)
                if not tipo.empty:
                    tipo["Aderência %"] = (tipo["Aderentes"] / tipo["Eventos"] * 100).round(1)
                    st.dataframe(
                        tipo.style.format({"Aderência %": "{:.1f}"}, na_rep="—")
                            .background_gradient(cmap="Purples", subset=["Eventos"]),
                        use_container_width=True, hide_index=True,
                    )
            with col_causa:
                st.markdown("**Por Causa**")
                causa = etit_por_causa(_etit_eq)
                if not causa.empty:
                    causa["Aderência %"] = (causa["Aderentes"] / causa["Eventos"] * 100).round(1)
                    st.dataframe(
                        causa.head(15).style.format({"Aderência %": "{:.1f}"}, na_rep="—")
                            .background_gradient(cmap="Purples", subset=["Eventos"]),
                        use_container_width=True, hide_index=True,
                    )

            col_reg, col_turno = st.columns(2)
            with col_reg:
                st.markdown("**Por Regional**")
                reg = etit_por_regional(_etit_eq)
                if not reg.empty:
                    reg["Aderência %"] = (reg["Aderentes"] / reg["Eventos"] * 100).round(1)
                    st.dataframe(reg.style.format({"Aderência %": "{:.1f}"}, na_rep="—"), use_container_width=True, hide_index=True)
            with col_turno:
                st.markdown("**Por Turno**")
                turno = etit_por_turno(_etit_eq)
                if not turno.empty:
                    turno["Aderência %"] = (turno["Aderentes"] / turno["Eventos"] * 100).round(1)
                    st.dataframe(turno.style.format({"Aderência %": "{:.1f}"}, na_rep="—"), use_container_width=True, hide_index=True)

            st.markdown("---")
            st.markdown("##### 📅 Evolução Diária ETIT")
            daily_etit = etit_evolucao_diaria(_etit_eq)
            if not daily_etit.empty:
                st.area_chart(daily_etit[["Data", "Eventos"]].set_index("Data"), color="#8E44AD", height=250)

            st.markdown("---")
            etit_show_cols = [
                ETIT_COL_LOGIN, "Nome", "Setor", ETIT_COL_DEMANDA, ETIT_COL_NOTA,
                ETIT_COL_STATUS, ETIT_COL_TIPO, ETIT_COL_CAUSA,
                ETIT_COL_GRUPO, ETIT_COL_REGIONAL, ETIT_COL_CIDADE, ETIT_COL_UF,
                ETIT_COL_TURNO, ETIT_COL_TMA, ETIT_COL_TMR,
                ETIT_COL_DT_ACIONAMENTO, ETIT_COL_ANOMES,
            ]
            etit_show_cols = [c for c in etit_show_cols if c in _etit_eq.columns]
            st.dataframe(
                _etit_eq[etit_show_cols].sort_values(
                    [ETIT_COL_DT_ACIONAMENTO] if ETIT_COL_DT_ACIONAMENTO in _etit_eq.columns else ["Nome"],
                    ascending=False,
                ),
                use_container_width=True, height=500,
            )
            csv_etit = _etit_eq[etit_show_cols].to_csv(index=False).encode("utf-8")
            st.download_button("📥 Baixar ETIT filtrado (CSV)", csv_etit, "etit_por_evento_equipe.csv", "text/csv")


# ---- TAB: INDICADORES RESIDENCIAL ----
if res_ind_loaded and _tab_res_idx is not None:
    with tabs[_tab_res_idx]:
        # ---- Filtrar somente regional Leste (equipe do usuário) ----
        _res_data = df_res_filtrado.copy()
        if RES_REGIONAL in _res_data.columns:
            _res_data = _res_data[_res_data[RES_REGIONAL].str.contains("Leste", case=False, na=False)]
        st.markdown("#### 🏠 Indicadores Residencial — Regional Leste")
        st.caption("🗺️ Dados filtrados para a regional **Leste** (equipe monitorada).")

        if _res_data.empty:
            st.warning("Nenhum dado encontrado para a regional Leste com os filtros atuais.")
        else:
            kpis_df = res_kpis_por_indicador(_res_data)
            st.markdown("##### 📊 Resumo por Indicador")
            n_cols = len(kpis_df)
            ind_cols = st.columns(n_cols) if n_cols > 0 else []
            for i, row in kpis_df.iterrows():
                ind_name = row["Indicador"]
                label = RES_IND_LABELS.get(ind_name, ind_name)
                color = RES_IND_COLORS.get(ind_name, "#5DADE2")
                vol = int(row["Volume"]); ader = int(row["Aderentes"]); pct = row["Aderencia_Pct"]
                tma_str = f"TMA: {row['TMA_Medio']:.4f}" if "TMA_Medio" in row and pd.notna(row.get("TMA_Medio")) else ""
                tmr_str = f"TMR: {row['TMR_Medio']:.4f}" if "TMR_Medio" in row and pd.notna(row.get("TMR_Medio")) else ""
                extra = " · ".join(filter(None, [tma_str, tmr_str]))
                pct_color = COR_SUCESSO if pct >= 90 else (COR_ALERTA if pct >= 70 else COR_PERIGO)
                with ind_cols[i]:
                    st.markdown(f"""<div class="res-ind-card" style="border-top-color:{color};">
                        <div class="ri-title">{label}</div>
                        <div class="ri-vol" style="color:{color};">{vol:,}</div>
                        <div class="ri-pct" style="color:{pct_color};">✅ {ader:,} aderentes &nbsp;·&nbsp; {pct:.1f}%</div>
                        <div class="ri-detail">{extra}</div>
                    </div>""", unsafe_allow_html=True)

            st.markdown("")
            st.markdown("##### 📈 Comparativo de Aderência por Indicador")
            if not kpis_df.empty:
                chart_ader = kpis_df[["Indicador", "Aderencia_Pct"]].copy()
                chart_ader["Indicador"] = chart_ader["Indicador"].map(RES_IND_LABELS)
                chart_ader = chart_ader.set_index("Indicador")
                chart_ader.columns = ["Aderência %"]
                st.bar_chart(chart_ader, color=COR_SUCESSO, height=300)

            # ---- IN_GRUPO breakdown ----
            if RES_GRUPO in _res_data.columns:
                st.markdown("---")
                st.markdown("##### 📍 Desempenho por Grupo (IN_GRUPO)")
                _rg = _res_data.groupby(RES_GRUPO).agg(
                    Volume=(RES_COL_VOLUME, "sum"), Aderentes=("ADERENTE", "sum"),
                ).reset_index().rename(columns={RES_GRUPO: "Grupo"})
                _rg["Aderência %"] = (_rg["Aderentes"] / _rg["Volume"] * 100).round(1)
                _rg = _rg.sort_values("Volume", ascending=False).reset_index(drop=True)
                if not _rg.empty:
                    _rg_best = _rg.loc[_rg["Aderência %"].idxmax()]
                    _rg_worst = _rg.loc[_rg["Aderência %"].idxmin()]
                    st.caption(
                        f"🟢 Melhor grupo: **{_rg_best['Grupo']}** ({_rg_best['Aderência %']:.1f}%) · "
                        f"🔴 Pior grupo: **{_rg_worst['Grupo']}** ({_rg_worst['Aderência %']:.1f}%)"
                    )
                    _srg = _rg.style.format({"Aderência %": "{:.1f}"}, na_rep="—")
                    _srg = _srg.background_gradient(cmap="RdYlGn", subset=["Aderência %"], vmin=50, vmax=100)
                    _srg = _srg.background_gradient(cmap="Blues", subset=["Volume"])
                    st.dataframe(_srg, use_container_width=True, hide_index=True)

                    # IN_GRUPO por indicador
                    st.markdown("**IN_GRUPO por Indicador**")
                    for _ri in RES_INDICADORES_FILTRO:
                        _ri_sub = _res_data[_res_data[RES_COL_INDICADOR_NOME] == _ri]
                        if _ri_sub.empty:
                            continue
                        _ri_g = _ri_sub.groupby(RES_GRUPO).agg(
                            Volume=(RES_COL_VOLUME, "sum"), Aderentes=("ADERENTE", "sum"),
                        ).reset_index().rename(columns={RES_GRUPO: "Grupo"})
                        _ri_g["Aderência %"] = (_ri_g["Aderentes"] / _ri_g["Volume"] * 100).round(1)
                        _ri_g = _ri_g.sort_values("Volume", ascending=False).reset_index(drop=True)
                        if not _ri_g.empty:
                            _ri_best = _ri_g.loc[_ri_g["Aderência %"].idxmax()]
                            _ri_worst = _ri_g.loc[_ri_g["Aderência %"].idxmin()]
                            with st.expander(
                                f"📊 {RES_IND_LABELS.get(_ri, _ri)} — "
                                f"Melhor: {_ri_best['Grupo']} ({_ri_best['Aderência %']:.1f}%) · "
                                f"Pior: {_ri_worst['Grupo']} ({_ri_worst['Aderência %']:.1f}%)"
                            ):
                                st.dataframe(
                                    _ri_g.style.format({"Aderência %": "{:.1f}"}, na_rep="—")
                                        .background_gradient(cmap="RdYlGn", subset=["Aderência %"], vmin=50, vmax=100),
                                    use_container_width=True, hide_index=True,
                                )

            st.markdown("---")
            ind_to_show = (
                [res_ind_selecionado] if res_ind_selecionado != "Todos" else RES_INDICADORES_FILTRO
            )
            ind_to_show = [i for i in ind_to_show if i in _res_data[RES_COL_INDICADOR_NOME].unique()]

            for ind in ind_to_show:
                label = RES_IND_LABELS.get(ind, ind)
                color = RES_IND_COLORS.get(ind, "#5DADE2")
                sub = _res_data[_res_data[RES_COL_INDICADOR_NOME] == ind]
                if sub.empty:
                    continue
                vol_total = int(sub[RES_COL_VOLUME].sum())
                ader_total = int(sub["ADERENTE"].sum())
                pct_total = (ader_total / vol_total * 100) if vol_total > 0 else 0

                with st.expander(f"🔍 {label} — {vol_total:,} registros · {pct_total:.1f}% aderência",
                                 expanded=(len(ind_to_show) == 1)):
                    sk1, sk2, sk3, sk4, sk5 = st.columns(5)
                    with sk1:
                        st.markdown(kpi_card("Volume", f"{vol_total:,}", color), unsafe_allow_html=True)
                    with sk2:
                        st.markdown(kpi_card("Aderentes", f"{ader_total:,}", COR_SUCESSO), unsafe_allow_html=True)
                    with sk3:
                        pct_c = COR_SUCESSO if pct_total >= 90 else (COR_ALERTA if pct_total >= 70 else COR_PERIGO)
                        st.markdown(kpi_card("Aderência", f"{pct_total:.1f}", pct_c, suffix="%"), unsafe_allow_html=True)
                    with sk4:
                        if RES_TMA in sub.columns:
                            st.markdown(kpi_card("TMA Médio", f"{sub[RES_TMA].mean():.4f}", COR_INFO), unsafe_allow_html=True)
                    with sk5:
                        if RES_TMR in sub.columns:
                            st.markdown(kpi_card("TMR Médio", f"{sub[RES_TMR].mean():.4f}", COR_ALERTA), unsafe_allow_html=True)

                    cn, cs = st.columns(2)
                    with cn:
                        st.markdown("**Por Natureza**")
                        nat_df = res_por_natureza(sub)
                        if not nat_df.empty:
                            st.dataframe(
                                nat_df.style.format({"Aderencia_Pct": "{:.1f}"}, na_rep="—")
                                    .background_gradient(cmap="RdYlGn", subset=["Aderencia_Pct"], vmin=50, vmax=100),
                                use_container_width=True, hide_index=True,
                            )
                        st.markdown("**Por Impacto**")
                        imp_df = res_por_impacto(sub)
                        if not imp_df.empty:
                            st.dataframe(imp_df.style.format({"Aderencia_Pct": "{:.1f}"}, na_rep="—"),
                                         use_container_width=True, hide_index=True)
                    with cs:
                        st.markdown("**Por Solução (Top 15)**")
                        sol_df = res_por_solucao(sub, top_n=15)
                        if not sol_df.empty:
                            st.dataframe(
                                sol_df.style.format({"Aderencia_Pct": "{:.1f}"}, na_rep="—")
                                    .background_gradient(cmap="YlOrRd", subset=["Volume"]),
                                use_container_width=True, hide_index=True, height=350,
                            )

                    st.markdown("**Evolução Diária**")
                    evo_df = res_evolucao_diaria(sub)
                    if not evo_df.empty:
                        c_evo1, c_evo2 = st.columns(2)
                        with c_evo1:
                            st.caption("Volume diário")
                            st.area_chart(evo_df[["Data", "Volume"]].set_index("Data"), color=color, height=200)
                        with c_evo2:
                            st.caption("Aderência diária (%)")
                            st.line_chart(evo_df[["Data", "Aderencia_Pct"]].set_index("Data"), color=COR_SUCESSO, height=200)

            st.markdown("---")

            # =============================================
            # INSIGHTS — INDICADORES RESIDENCIAL (LESTE)
            # =============================================
            st.markdown("### 💡 Insights — Indicadores Residencial (Leste)")
            _res_insights = []

            if not kpis_df.empty:
                # Melhor e pior indicador
                best_ind = kpis_df.loc[kpis_df["Aderencia_Pct"].idxmax()]
                worst_ind = kpis_df.loc[kpis_df["Aderencia_Pct"].idxmin()]
                _res_insights.append(
                    f"🟢 Melhor indicador: **{RES_IND_LABELS.get(best_ind['Indicador'], best_ind['Indicador'])}** "
                    f"com **{best_ind['Aderencia_Pct']:.1f}%** de aderência ({int(best_ind['Volume'])} registros)."
                )
                if best_ind["Indicador"] != worst_ind["Indicador"]:
                    _res_insights.append(
                        f"🔴 Pior indicador: **{RES_IND_LABELS.get(worst_ind['Indicador'], worst_ind['Indicador'])}** "
                        f"com **{worst_ind['Aderencia_Pct']:.1f}%** de aderência ({int(worst_ind['Volume'])} registros)."
                    )
                # Média geral
                _avg_ader = kpis_df["Aderencia_Pct"].mean()
                _res_insights.append(f"📊 Aderência média geral (Leste): **{_avg_ader:.1f}%**.")

                # Indicadores abaixo de 80%
                _low = kpis_df[kpis_df["Aderencia_Pct"] < 80]
                if len(_low) > 0:
                    _low_names = [RES_IND_LABELS.get(r["Indicador"], r["Indicador"]) for _, r in _low.iterrows()]
                    _res_insights.append(f"⚠️ {len(_low)} indicador(es) abaixo de 80%: {', '.join(_low_names)}.")

                # Volume total
                _vol_total = int(kpis_df["Volume"].sum())
                _ader_total = int(kpis_df["Aderentes"].sum())
                _res_insights.append(f"📦 Volume total na Leste: **{_vol_total:,}** registros, **{_ader_total:,}** aderentes.")

                # Por natureza — top causa de não aderência
                for ind in RES_INDICADORES_FILTRO:
                    sub_i = _res_data[_res_data[RES_COL_INDICADOR_NOME] == ind]
                    if sub_i.empty:
                        continue
                    nat_i = res_por_natureza(sub_i)
                    if not nat_i.empty:
                        worst_nat = nat_i.loc[nat_i["Aderencia_Pct"].idxmin()]
                        if worst_nat["Aderencia_Pct"] < 70:
                            _res_insights.append(
                                f"🔍 Em **{RES_IND_LABELS.get(ind, ind)}**, a natureza "
                                f"**{worst_nat['Natureza']}** tem apenas **{worst_nat['Aderencia_Pct']:.1f}%** de aderência."
                            )

            if _res_insights:
                for ins in _res_insights:
                    st.markdown(ins)
            else:
                st.caption("Sem dados suficientes para gerar insights.")

            st.markdown("---")
            res_show_cols = [
                RES_COL_INDICADOR_NOME, RES_COL_ID_MOSTRA, RES_COL_VOLUME,
                "ADERENTE", RES_COL_STATUS, RES_REGIONAL, RES_COL_NATUREZA,
                RES_COL_IMPACTO, RES_COL_SOLUCAO, RES_TMA, RES_TMR,
                RES_COL_DT_INICIO, RES_ANOMES,
            ]
            res_show_cols = [c for c in res_show_cols if c in _res_data.columns]
            st.dataframe(
                _res_data[res_show_cols].sort_values(
                    [RES_COL_DT_INICIO] if RES_COL_DT_INICIO in _res_data.columns else [RES_COL_INDICADOR_NOME],
                    ascending=False,
                ),
                use_container_width=True, height=400,
            )
            csv_res = _res_data[res_show_cols].to_csv(index=False).encode("utf-8")
            st.download_button("📥 Baixar Indicadores Residencial Leste (CSV)", csv_res, "indicadores_residencial_leste.csv", "text/csv")



# ---- TAB: INDICADORES TOA ----
if toa_loaded and _tab_toa_idx is not None:
    with tabs[_tab_toa_idx]:
        # ---- Filtrar somente EMPRESARIAL e excluir líderes ----
        df_toa_emp = df_toa[(df_toa["Setor"] == "EMPRESARIAL") & (~df_toa["LOGIN"].isin(LIDERES_IDS))].copy()
        anomes_str = str(toa_anomes) if toa_anomes else "?"
        st.markdown(
            f"#### 📋 Indicadores TOA — Empresarial · "
            f"Período: **{anomes_str}** (mês mais recente)"
        )
        st.caption(
            "🏢 Dados dos analistas empresariais (sem líderes). "
            "Tarefas Canceladas: menor = melhor. Tempo de Validação: maior aderência% = melhor."
        )

        if df_toa_emp.empty:
            st.warning("Nenhum dado TOA encontrado para o setor Empresarial.")
        else:
            # ---- KPIs gerais ----
            resumo_toa = toa_resumo_por_indicador(df_toa_emp)
            if not resumo_toa.empty:
                tk_cols = st.columns(len(resumo_toa) * 2)
                ci = 0
                for _, trow in resumo_toa.iterrows():
                    ind_nome = trow["Indicador"]
                    cor = TOA_IND_COLORS.get(ind_nome, COR_INFO)
                    label = TOA_IND_LABELS.get(ind_nome, ind_nome)
                    with tk_cols[ci]:
                        st.markdown(kpi_card(f"Total — {label[:20]}", f"{trow['Total']:,}", cor), unsafe_allow_html=True)
                    ci += 1
                    with tk_cols[ci]:
                        if ind_nome == TOA_IND_CANCELADAS:
                            st.markdown(kpi_card("Canceladas (⚠️ menor melhor)", f"{trow['Total']:,}", COR_PERIGO), unsafe_allow_html=True)
                        else:
                            pct = trow["Aderencia_Pct"]
                            pct_c = COR_SUCESSO if pct >= 90 else (COR_ALERTA if pct >= 70 else COR_PERIGO)
                            st.markdown(kpi_card(f"Aderência — {label[:15]}", f"{pct:.1f}", pct_c, suffix="%"), unsafe_allow_html=True)
                    ci += 1

            st.markdown("---")

            # =============================================
            # SEÇÃO 1: TAREFAS CANCELADAS
            # =============================================
            st.markdown("### ❌ Tarefas Canceladas")
            st.caption("Cada linha representa uma tarefa cancelada por um analista empresarial no período.")

            col_canc1, col_canc2 = st.columns([1, 1])

            with col_canc1:
                st.markdown("##### 🏆 Ranking por Analista")
                df_canc_anal = toa_canceladas_por_analista(df_toa_emp)
                if not df_canc_anal.empty:
                    df_canc_anal["Analista"] = df_canc_anal["Nome"].apply(primeiro_nome)
                    df_canc_anal["TMR Médio (h)"] = df_canc_anal["TMR_Medio_h"]
                    tbl_canc = df_canc_anal[["Analista", "Canceladas", "TMR Médio (h)"]].copy()
                    tbl_canc = tbl_canc.reset_index(drop=True)
                    tbl_canc.index += 1; tbl_canc.index.name = "#"
                    st.dataframe(
                        tbl_canc.style
                            .format({"TMR Médio (h)": "{:.2f}"}, na_rep="—")
                            .background_gradient(cmap="Reds", subset=["Canceladas"]),
                        use_container_width=True,
                    )
                    chart_canc = df_canc_anal[["Analista", "Canceladas"]].set_index("Analista").sort_values("Canceladas")
                    st.bar_chart(chart_canc, color="#E74C3C", height=300)

            with col_canc2:
                st.markdown("##### ⏱️ Distribuição por Faixa de Tempo (AGING)")
                df_aging = toa_canceladas_por_aging(df_toa_emp)
                if not df_aging.empty:
                    st.dataframe(
                        df_aging.style.background_gradient(cmap="Reds", subset=["Canceladas"]),
                        use_container_width=True, hide_index=True,
                    )
                    chart_aging = df_aging.set_index("Aging")
                    st.bar_chart(chart_aging, color="#E74C3C", height=250)

                st.markdown("##### 🔧 Por Tipo de Atividade")
                df_canc_tipo = toa_canceladas_por_tipo(df_toa_emp)
                if not df_canc_tipo.empty:
                    st.dataframe(
                        df_canc_tipo.style.background_gradient(cmap="Reds", subset=["Canceladas"]),
                        use_container_width=True, hide_index=True,
                    )

            col_cr, col_creg = st.columns(2)
            with col_cr:
                st.markdown("##### 📡 Por Rede")
                df_canc_rede = toa_canceladas_por_rede(df_toa_emp)
                if not df_canc_rede.empty:
                    st.dataframe(
                        df_canc_rede.style.background_gradient(cmap="Reds", subset=["Canceladas"]),
                        use_container_width=True, hide_index=True,
                    )
            with col_creg:
                st.markdown("##### 🗺️ Por Regional")
                df_canc_reg = toa_canceladas_por_regional(df_toa_emp)
                if not df_canc_reg.empty:
                    st.dataframe(
                        df_canc_reg.style.background_gradient(cmap="Reds", subset=["Canceladas"]),
                        use_container_width=True, hide_index=True,
                    )

            # ---- SOLUCAO para Canceladas ----
            _canc_sub = df_toa_emp[df_toa_emp["INDICADOR_NOME"] == TOA_IND_CANCELADAS]
            col_sol, col_grp = st.columns(2)
            with col_sol:
                if "SOLUCAO" in _canc_sub.columns:
                    st.markdown("##### 🔧 Por Solução")
                    _cs = _canc_sub.groupby("SOLUCAO").size().reset_index(name="Canceladas")
                    _cs = _cs.sort_values("Canceladas", ascending=False).reset_index(drop=True)
                    if not _cs.empty:
                        st.dataframe(
                            _cs.head(15).style.background_gradient(cmap="Reds", subset=["Canceladas"]),
                            use_container_width=True, hide_index=True,
                        )
            with col_grp:
                if "IN_GRUPO" in _canc_sub.columns:
                    st.markdown("##### 📍 Canceladas por Grupo (IN_GRUPO)")
                    _cg = _canc_sub.groupby("IN_GRUPO").size().reset_index(name="Canceladas")
                    _cg = _cg.sort_values("Canceladas", ascending=False).reset_index(drop=True)
                    _cg["% do Total"] = (_cg["Canceladas"] / _cg["Canceladas"].sum() * 100).round(1)
                    if not _cg.empty:
                        _cg_best = _cg.loc[_cg["Canceladas"].idxmin()]
                        _cg_worst = _cg.loc[_cg["Canceladas"].idxmax()]
                        st.caption(
                            f"🟢 Menor: **{_cg_best['IN_GRUPO']}** ({int(_cg_best['Canceladas'])}) · "
                            f"🔴 Maior: **{_cg_worst['IN_GRUPO']}** ({int(_cg_worst['Canceladas'])})"
                        )
                        st.dataframe(
                            _cg.style.background_gradient(cmap="Reds", subset=["Canceladas"]),
                            use_container_width=True, hide_index=True,
                        )

            st.markdown("##### 📅 Evolução Diária — Canceladas")
            df_canc_evo = toa_canceladas_evolucao(df_toa_emp)
            if not df_canc_evo.empty:
                st.area_chart(df_canc_evo[["Data", "Canceladas"]].set_index("Data"), color="#E74C3C", height=220)

            st.markdown("---")

            # =============================================
            # SEÇÃO 2: TEMPO DE VALIDAÇÃO DO FORMULÁRIO
            # =============================================
            st.markdown("### ✅ Tempo de Validação do Formulário")
            st.caption("Aderência ao tempo máximo permitido para validar o formulário TOA. Maior aderência% = melhor.")

            col_val1, col_val2 = st.columns([1, 1])

            with col_val1:
                st.markdown("##### 🏆 Ranking por Analista")
                df_val_anal = toa_validacao_por_analista(df_toa_emp)
                if not df_val_anal.empty:
                    df_val_anal["Analista"] = df_val_anal["Nome"].apply(primeiro_nome)
                    tbl_val = df_val_anal[["Analista", "Total", "Aderentes", "Aderencia_Pct", "TMR_Medio_min"]].copy()
                    tbl_val.columns = ["Analista", "Total", "Aderentes", "Aderência %", "TMR Médio (min)"]
                    tbl_val = tbl_val.reset_index(drop=True)
                    tbl_val.index += 1; tbl_val.index.name = "#"
                    styled_val = tbl_val.style.format(
                        {"Aderência %": "{:.1f}", "TMR Médio (min)": "{:.1f}"}, na_rep="—"
                    )
                    styled_val = styled_val.background_gradient(cmap="RdYlGn", subset=["Aderência %"], vmin=40, vmax=100)
                    styled_val = styled_val.background_gradient(cmap="RdYlGn_r", subset=["TMR Médio (min)"], vmin=5, vmax=60)
                    st.dataframe(styled_val, use_container_width=True)

                    if len(tbl_val) >= 2:
                        best_v = tbl_val.iloc[0]
                        worst_v = tbl_val.iloc[-1]
                        cv1, cv2 = st.columns(2)
                        with cv1:
                            st.markdown(f"""<div class="perf-card perf-best">
                                <div class="p-title">🏆 Melhor Aderência</div>
                                <div class="p-name" style="color:#2ECC71;">{best_v['Analista']}</div>
                                <div class="p-detail">{best_v['Aderência %']:.1f}% · TMR: {best_v['TMR Médio (min)']:.1f} min</div>
                            </div>""", unsafe_allow_html=True)
                        with cv2:
                            st.markdown(f"""<div class="perf-card perf-worst">
                                <div class="p-title">⚠️ Menor Aderência</div>
                                <div class="p-name" style="color:#E74C3C;">{worst_v['Analista']}</div>
                                <div class="p-detail">{worst_v['Aderência %']:.1f}% · TMR: {worst_v['TMR Médio (min)']:.1f} min</div>
                            </div>""", unsafe_allow_html=True)

            with col_val2:
                st.markdown("##### 📊 Aderência por Analista")
                if not df_val_anal.empty:
                    chart_val = df_val_anal[["Analista", "Aderencia_Pct"]].set_index("Analista").sort_values("Aderencia_Pct")
                    chart_val.columns = ["Aderência %"]
                    st.bar_chart(chart_val, horizontal=True, color="#16A085", height=400)

            col_v1, col_v2, col_v3 = st.columns(3)
            with col_v1:
                st.markdown("##### 🔧 Por Tipo de Atividade")
                df_val_tipo = toa_validacao_por_tipo(df_toa_emp)
                if not df_val_tipo.empty:
                    df_val_tipo_show = df_val_tipo[["Tipo Atividade", "Total", "Aderentes", "Aderencia_Pct", "TMR_Medio_min"]].copy()
                    df_val_tipo_show.columns = ["Tipo Atividade", "Total", "Aderentes", "Aderência %", "TMR (min)"]
                    st.dataframe(
                        df_val_tipo_show.style
                            .format({"Aderência %": "{:.1f}", "TMR (min)": "{:.1f}"}, na_rep="—")
                            .background_gradient(cmap="RdYlGn", subset=["Aderência %"], vmin=40, vmax=100),
                        use_container_width=True, hide_index=True,
                    )
            with col_v2:
                st.markdown("##### 📡 Por Rede")
                df_val_rede = toa_validacao_por_rede(df_toa_emp)
                if not df_val_rede.empty:
                    df_val_rede_show = df_val_rede[["Rede", "Total", "Aderentes", "Aderencia_Pct", "TMR_Medio_min"]].copy()
                    df_val_rede_show.columns = ["Rede", "Total", "Aderentes", "Aderência %", "TMR (min)"]
                    st.dataframe(
                        df_val_rede_show.style
                            .format({"Aderência %": "{:.1f}", "TMR (min)": "{:.1f}"}, na_rep="—")
                            .background_gradient(cmap="RdYlGn", subset=["Aderência %"], vmin=40, vmax=100),
                        use_container_width=True, hide_index=True,
                    )
            with col_v3:
                st.markdown("##### 🗺️ Por Regional")
                df_val_reg = toa_validacao_por_regional(df_toa_emp)
                if not df_val_reg.empty:
                    df_val_reg_show = df_val_reg[["Regional", "Total", "Aderentes", "Aderencia_Pct", "TMR_Medio_min"]].copy()
                    df_val_reg_show.columns = ["Regional", "Total", "Aderentes", "Aderência %", "TMR (min)"]
                    st.dataframe(
                        df_val_reg_show.style
                            .format({"Aderência %": "{:.1f}", "TMR (min)": "{:.1f}"}, na_rep="—")
                            .background_gradient(cmap="RdYlGn", subset=["Aderência %"], vmin=40, vmax=100),
                        use_container_width=True, hide_index=True,
                    )

            # ---- IN_GRUPO para Validação ----
            _val_sub = df_toa_emp[df_toa_emp["INDICADOR_NOME"] == TOA_IND_VALIDACAO]
            if "IN_GRUPO" in _val_sub.columns and not _val_sub.empty:
                st.markdown("##### 📍 Validação por Grupo (IN_GRUPO)")
                _vg = _val_sub.groupby("IN_GRUPO").agg(
                    Total=("INDICADOR", "count"),
                    Aderentes=("ADERENTE", "sum"),
                ).reset_index()
                _vg["Aderência %"] = (_vg["Aderentes"] / _vg["Total"] * 100).round(1)
                if "TMR_min" in _val_sub.columns:
                    _vg_tmr = _val_sub.groupby("IN_GRUPO")["TMR_min"].mean().reset_index()
                    _vg_tmr.columns = ["IN_GRUPO", "TMR (min)"]
                    _vg = _vg.merge(_vg_tmr, on="IN_GRUPO", how="left")
                _vg = _vg.sort_values("Aderência %", ascending=False).reset_index(drop=True)
                if not _vg.empty and len(_vg) >= 2:
                    _vg_best = _vg.iloc[0]
                    _vg_worst = _vg.iloc[-1]
                    st.caption(
                        f"🟢 Melhor: **{_vg_best['IN_GRUPO']}** ({_vg_best['Aderência %']:.1f}%) · "
                        f"🔴 Pior: **{_vg_worst['IN_GRUPO']}** ({_vg_worst['Aderência %']:.1f}%)"
                    )
                _fmt_vg = {"Aderência %": "{:.1f}"}
                if "TMR (min)" in _vg.columns:
                    _fmt_vg["TMR (min)"] = "{:.1f}"
                st.dataframe(
                    _vg.style
                        .format(_fmt_vg, na_rep="—")
                        .background_gradient(cmap="RdYlGn", subset=["Aderência %"], vmin=40, vmax=100),
                    use_container_width=True, hide_index=True,
                )

            st.markdown("##### 📅 Evolução Diária — Validação do Formulário")
            df_val_evo = toa_validacao_evolucao(df_toa_emp)
            if not df_val_evo.empty:
                c_evo1, c_evo2 = st.columns(2)
                with c_evo1:
                    st.caption("Aderência diária (%)")
                    st.line_chart(df_val_evo[["Data", "Aderencia_Pct"]].set_index("Data"), color="#16A085", height=220)
                with c_evo2:
                    st.caption("TMR médio diário (min)")
                    st.line_chart(df_val_evo[["Data", "TMR_Medio_min"]].set_index("Data"), color=COR_ALERTA, height=220)

            st.markdown("---")

            # =============================================
            # INSIGHTS TOA — EMPRESARIAL
            # =============================================
            st.markdown("### 💡 Insights — Indicadores TOA Empresarial")
            _toa_insights = []

            # Insights de canceladas
            _canc_a = toa_canceladas_por_analista(df_toa_emp)
            if not _canc_a.empty:
                _canc_a["Analista"] = _canc_a["Nome"].apply(primeiro_nome)
                _worst_canc = _canc_a.iloc[0]  # mais canceladas (sorted desc by default)
                _best_canc = _canc_a.iloc[-1]   # menos canceladas
                _total_canc = int(_canc_a["Canceladas"].sum())
                _media_canc = _canc_a["Canceladas"].mean()
                _toa_insights.append(f"🔴 **{_worst_canc['Analista']}** lidera em cancelamentos com **{int(_worst_canc['Canceladas'])}** tarefas ({_worst_canc['Canceladas']/_total_canc*100:.0f}% do total da equipe).")
                _toa_insights.append(f"🟢 **{_best_canc['Analista']}** tem o menor volume de cancelamentos: **{int(_best_canc['Canceladas'])}** tarefas.")
                _above_avg = _canc_a[_canc_a["Canceladas"] > _media_canc]
                if len(_above_avg) > 0:
                    _toa_insights.append(f"⚠️ {len(_above_avg)} analista(s) acima da média de {_media_canc:.0f} cancelamentos: {', '.join(_above_avg['Analista'].tolist())}.")

            # Insights de validação
            _val_a = toa_validacao_por_analista(df_toa_emp)
            if not _val_a.empty:
                _val_a["Analista"] = _val_a["Nome"].apply(primeiro_nome)
                _media_ader = _val_a["Aderencia_Pct"].mean()
                _best_val = _val_a.iloc[0]
                _worst_val = _val_a.iloc[-1]
                _toa_insights.append(f"📊 Aderência média da equipe na validação: **{_media_ader:.1f}%**.")
                _toa_insights.append(f"🏆 **{_best_val['Analista']}** tem a melhor aderência: **{_best_val['Aderencia_Pct']:.1f}%** (TMR: {_best_val['TMR_Medio_min']:.1f} min).")
                _toa_insights.append(f"⚠️ **{_worst_val['Analista']}** tem a menor aderência: **{_worst_val['Aderencia_Pct']:.1f}%** (TMR: {_worst_val['TMR_Medio_min']:.1f} min).")
                _low_ader = _val_a[_val_a["Aderencia_Pct"] < 80]
                if len(_low_ader) > 0:
                    _toa_insights.append(f"🔴 {len(_low_ader)} analista(s) com aderência abaixo de 80%: {', '.join(_low_ader['Analista'].tolist())}.")

            # Insights de tipo/rede
            _canc_t = toa_canceladas_por_tipo(df_toa_emp)
            if not _canc_t.empty and len(_canc_t) >= 1:
                _top_tipo = _canc_t.iloc[0]
                _toa_insights.append(f"🔧 Tipo de atividade com mais cancelamentos: **{_top_tipo['Tipo Atividade']}** ({int(_top_tipo['Canceladas'])} tarefas).")

            _canc_r = toa_canceladas_por_rede(df_toa_emp)
            if not _canc_r.empty and len(_canc_r) >= 1:
                _top_rede = _canc_r.iloc[0]
                _toa_insights.append(f"📡 Rede com mais cancelamentos: **{_top_rede['Rede']}** ({int(_top_rede['Canceladas'])} tarefas).")

            if _toa_insights:
                for ins in _toa_insights:
                    st.markdown(ins)
            else:
                st.caption("Sem dados suficientes para gerar insights.")

            st.markdown("---")

            # ---- Export ----
            st.markdown("##### 📥 Exportar dados TOA (Empresarial)")
            toa_export_cols = [
                "INDICADOR_NOME", "ID_ATIVIDADE", "LOGIN", "Nome", "Setor",
                "IN_REGIONAL", "TIPO_ATIVIDADE", "REDE", "MERCADO", "NATUREZA",
                "INDICADOR", "INDICADOR_STATUS", "ADERENTE",
                "TMR_min", "AGING", "DATA", "DT_CANCELAMENTO",
                "DT_INICIO_FORM", "DT_FIM_FORM", "ANOMES",
            ]
            toa_export_cols = [c for c in toa_export_cols if c in df_toa_emp.columns]
            csv_toa = df_toa_emp[toa_export_cols].to_csv(index=False).encode("utf-8")
            st.download_button(
                "📥 Baixar Indicadores TOA Empresarial (CSV)",
                csv_toa,
                f"indicadores_toa_empresarial_{anomes_str}.csv",
                "text/csv",
            )


# ---- TAB: OCUPAÇÃO DPA ----
if dpa_loaded and _tab_dpa_idx is not None:
    with tabs[_tab_dpa_idx]:
        mes_nome_dpa  = dpa_mes_info.get("mes_nome", "—")
        mes_num_dpa   = dpa_mes_info.get("mes_num")
        dpa_geral_pct = dpa_mes_info.get("dpa_geral_pct")

        st.markdown(
            f"#### 📊 Ocupação DPA — Dados Oficiais · "
            f"Mês mais recente: **{mes_nome_dpa} 2026**"
        )

        if mes_num_dpa:
            st.caption(
                f"ℹ️ O mês mais recente com dados disponíveis na planilha é **{mes_nome_dpa}**. "
                f"Os percentuais refletem a ocupação acumulada de Janeiro a {mes_nome_dpa} de 2026."
            )

        # KPIs gerais
        k1, k2, k3, k4 = st.columns(4)
        with k1:
            dpa_g_str = f"{dpa_geral_pct:.1f}" if dpa_geral_pct else "—"
            st.markdown(kpi_card(f"DPA Equipe ({mes_nome_dpa[:3]})", dpa_g_str, _dpa_color(dpa_geral_pct), suffix="%"),
                        unsafe_allow_html=True)
        with k2:
            st.markdown(kpi_card("Analistas Monitorados", str(len(df_dpa_filtrado)), COR_INFO), unsafe_allow_html=True)
        with k3:
            above = (df_dpa_filtrado["DPA_Pct_Oficial"] >= DPA_THRESHOLD_OK).sum()
            st.markdown(kpi_card(f"Acima de {DPA_THRESHOLD_OK:.0f}% 🟢", str(above), COR_SUCESSO), unsafe_allow_html=True)
        with k4:
            below = (df_dpa_filtrado["DPA_Pct_Oficial"] < DPA_THRESHOLD_ALERTA).sum()
            st.markdown(kpi_card(f"Abaixo de {DPA_THRESHOLD_ALERTA:.0f}% 🔴", str(below), COR_PERIGO), unsafe_allow_html=True)

        st.markdown("")
        st.markdown("---")

        # ---- Ranking principal ----
        st.markdown("##### 🏆 Ranking de Ocupação DPA por Analista")
        rank_of = dpa_ranking(df_dpa_filtrado)
        if not rank_of.empty:
            rank_of = rank_of[~rank_of["Login"].isin(LIDERES_IDS)].reset_index(drop=True)
            rank_of.index += 1; rank_of.index.name = "#"
            rank_of["Status"] = rank_of["DPA %"].apply(_dpa_semaforo)
            rank_of_display = rank_of[["Status", "Analista", "Setor", "DPA %"]]

            col_tbl, col_chart = st.columns([1, 1])
            with col_tbl:
                st.dataframe(
                    rank_of_display.style
                        .format({"DPA %": "{:.1f}"})
                        .background_gradient(cmap="RdYlGn", subset=["DPA %"], vmin=50, vmax=100),
                    use_container_width=True, height=560,
                )
            with col_chart:
                chart_dpa = rank_of[["Analista", "DPA %"]].set_index("Analista").sort_values("DPA %")
                st.bar_chart(chart_dpa, color="#16A085", height=560)

        st.markdown("---")

        # ---- Breakdown por Setor ----
        st.markdown("##### 🏢🏠 Ocupação DPA por Setor")
        c_emp, c_res = st.columns(2)
        for col_s, setor_s, cmap_s in [
            (c_emp, "EMPRESARIAL", "Oranges"),
            (c_res, "RESIDENCIAL", "Blues"),
        ]:
            with col_s:
                icon_s = "🏢" if setor_s == "EMPRESARIAL" else "🏠"
                st.markdown(f"**{icon_s} {setor_s}**")
                df_sec_s = df_dpa_filtrado[(df_dpa_filtrado["Setor"] == setor_s) & (~df_dpa_filtrado["Login"].isin(LIDERES_IDS))].copy()
                if df_sec_s.empty:
                    st.caption("Sem dados.")
                    continue
                media_s = df_sec_s["DPA_Pct_Oficial"].mean()
                df_sec_s["Nome_Curto"] = df_sec_s["Nome"].apply(primeiro_nome)
                df_sec_s["Status"] = df_sec_s["DPA_Pct_Oficial"].apply(_dpa_semaforo)
                df_sec_s_show = df_sec_s[["Status", "Nome_Curto", "DPA_Pct_Oficial"]].copy()
                df_sec_s_show.columns = ["Status", "Analista", "DPA %"]
                df_sec_s_show = df_sec_s_show.sort_values("DPA %", ascending=False).reset_index(drop=True)
                df_sec_s_show.index += 1; df_sec_s_show.index.name = "#"
                st.caption(f"Média do setor: **{media_s:.1f}%** {_dpa_semaforo(media_s)}")
                st.dataframe(
                    df_sec_s_show.style
                        .format({"DPA %": "{:.1f}"})
                        .background_gradient(cmap="RdYlGn", subset=["DPA %"], vmin=50, vmax=100),
                    use_container_width=True, height=400,
                )

        st.markdown("---")

        # ---- Semáforos visuais ----
        st.markdown("##### 🚦 Painel de Semáforo — Todos os Analistas")
        n_cards = 4
        card_cols = st.columns(n_cards)
        sorted_dpa = df_dpa_filtrado[~df_dpa_filtrado["Login"].isin(LIDERES_IDS)].sort_values("DPA_Pct_Oficial", ascending=False).reset_index(drop=True)
        for ci, (_, arow) in enumerate(sorted_dpa.iterrows()):
            pct_v = arow["DPA_Pct_Oficial"]
            nome_c = primeiro_nome(arow["Nome"])
            setor_c = arow["Setor"][:3]
            sem_icon = _dpa_semaforo(pct_v)
            sem_color = (
                "#27AE60" if pct_v >= DPA_THRESHOLD_OK
                else "#F39C12" if pct_v >= DPA_THRESHOLD_ALERTA
                else "#E74C3C"
            )
            with card_cols[ci % n_cards]:
                st.markdown(f"""<div class="dpa-card" style="border-left-color:{sem_color};">
                    <div class="dpa-nome">{sem_icon} {nome_c} <span style="font-size:0.72rem;opacity:0.55;">{setor_c}</span></div>
                    <div class="dpa-val" style="color:{sem_color};">{pct_v:.1f}%</div>
                </div>""", unsafe_allow_html=True)

        st.markdown("---")

        # =============================================
        # INSIGHTS — OCUPAÇÃO DPA
        # =============================================
        st.markdown("### 💡 Insights — Ocupação DPA")
        _dpa_insights = []

        if not df_dpa_filtrado.empty:
            _avg_of = df_dpa_filtrado["DPA_Pct_Oficial"].mean()
            _best_dpa = df_dpa_filtrado.loc[df_dpa_filtrado["DPA_Pct_Oficial"].idxmax()]
            _worst_dpa = df_dpa_filtrado.loc[df_dpa_filtrado["DPA_Pct_Oficial"].idxmin()]
            _dpa_insights.append(f"📊 DPA médio oficial da equipe: **{_avg_of:.1f}%** {_dpa_semaforo(_avg_of)}.")
            _dpa_insights.append(
                f"🏆 Maior DPA: **{primeiro_nome(_best_dpa['Nome'])}** com **{_best_dpa['DPA_Pct_Oficial']:.1f}%** "
                f"({_best_dpa['Setor'][:3]})."
            )
            _dpa_insights.append(
                f"⚠️ Menor DPA: **{primeiro_nome(_worst_dpa['Nome'])}** com **{_worst_dpa['DPA_Pct_Oficial']:.1f}%** "
                f"({_worst_dpa['Setor'][:3]})."
            )

            _n_green = (df_dpa_filtrado["DPA_Pct_Oficial"] >= DPA_THRESHOLD_OK).sum()
            _n_red = (df_dpa_filtrado["DPA_Pct_Oficial"] < DPA_THRESHOLD_ALERTA).sum()
            _n_total = len(df_dpa_filtrado)
            _dpa_insights.append(f"🟢 {_n_green}/{_n_total} analistas acima de {DPA_THRESHOLD_OK:.0f}%.")
            if _n_red > 0:
                _red_names = df_dpa_filtrado[df_dpa_filtrado["DPA_Pct_Oficial"] < DPA_THRESHOLD_ALERTA]["Nome"].apply(primeiro_nome).tolist()
                _dpa_insights.append(f"🔴 {_n_red} analista(s) abaixo de {DPA_THRESHOLD_ALERTA:.0f}%: {', '.join(_red_names)}.")

            # Comparação por setor
            for _s in ["EMPRESARIAL", "RESIDENCIAL"]:
                _sec = df_dpa_filtrado[df_dpa_filtrado["Setor"] == _s]
                if not _sec.empty:
                    _icon = "🏢" if _s == "EMPRESARIAL" else "🏠"
                    _dpa_insights.append(f"{_icon} {_s}: média **{_sec['DPA_Pct_Oficial'].mean():.1f}%** ({len(_sec)} analistas).")

        if _dpa_insights:
            for ins in _dpa_insights:
                st.markdown(ins)
        else:
            st.caption("Sem dados suficientes para gerar insights.")

        # ---- Export ----
        st.markdown("---")
        csv_dpa = df_dpa_filtrado[["Login", "Nome", "Setor", "DPA_Pct_Oficial"]].copy()
        csv_dpa.columns = ["Login", "Nome", "Setor", "DPA % Oficial"]
        st.download_button(
            "📥 Baixar Ocupação DPA (CSV)",
            csv_dpa.to_csv(index=False).encode("utf-8"),
            f"ocupacao_dpa_{mes_nome_dpa.lower()}_2026.csv",
            "text/csv",
        )



# =====================================================
# FOOTER
# =====================================================
st.markdown("---")
footer_parts = [
    "Dashboard de Produtividade COP Rede",
    f"{len(df_filtrado)} registros",
    f"{n_analistas} analistas",
    f"Dados de {data_min.strftime('%d/%m/%Y') if pd.notna(data_min) else '?'} "
    f"a {data_max.strftime('%d/%m/%Y') if pd.notna(data_max) else '?'}",
]
if etit_loaded:
    footer_parts.append(f"ETIT: {len(df_etit_filtrado)} eventos")
if res_ind_loaded:
    footer_parts.append(f"Ind. Residencial: {len(df_res_filtrado):,} registros")
if toa_loaded:
    footer_parts.append(f"TOA {toa_anomes}: {len(df_toa)} registros")
if dpa_loaded:
    footer_parts.append(f"DPA Oficial: {len(df_dpa_filtrado)} analistas · {dpa_mes_info.get('mes_nome','?')} 2026")
st.caption(" · ".join(footer_parts))
