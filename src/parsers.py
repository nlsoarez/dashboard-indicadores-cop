from __future__ import annotations

from dataclasses import dataclass
from typing import Optional, Iterable

import pandas as pd

from .config import (
    CHAT_TOA_NAME_PATTERNS,
    RES_ETIT_GPON_INDICADOR_NOME,
    RES_LOG_REPROG_GPON_INDICADOR_NOME,
    RES_SINTOMA_SEM_SINAL,
)

@dataclass
class IndicadorResult:
    indicador: str
    df: pd.DataFrame  # colunas: Matricula, aderente, total, pct

def _safe_upper(s: pd.Series) -> pd.Series:
    return s.astype(str).str.upper().str.strip()

def _calc_pct(aderente: pd.Series, total: pd.Series) -> pd.Series:
    return (aderente / total) * 100

def parse_chat_toa_from_toa(planilha_toa_xlsx, equipe_ids: set[str]) -> Optional[IndicadorResult]:
    """
    Tenta extrair o indicador 'Chat TOA' da planilha 'Analitico Indicadores TOA' (aba TOA).
    Nem todos os meses vem esse indicador nessa planilha; por isso é 'best-effort'.
    Esperado: colunas INDICADOR_NOME, INDICADOR (0/1), LOGIN.
    """
    df = pd.read_excel(planilha_toa_xlsx, sheet_name="TOA")
    required = {"INDICADOR_NOME", "INDICADOR", "LOGIN"}
    if not required.issubset(set(df.columns)):
        return None

    nome = _safe_upper(df["INDICADOR_NOME"])
    mask = pd.Series(False, index=df.index)
    for pat in CHAT_TOA_NAME_PATTERNS:
        mask = mask | nome.str.contains(pat, na=False)

    sub = df[mask].copy()
    if sub.empty:
        return None

    sub = sub[sub["LOGIN"].isin(equipe_ids)]
    if sub.empty:
        return None

    g = sub.groupby("LOGIN").agg(total=("INDICADOR","size"), aderente=("INDICADOR","sum")).reset_index()
    g["pct"] = _calc_pct(g["aderente"], g["total"])
    g = g.rename(columns={"LOGIN":"Matricula"})
    return IndicadorResult("Chat TOA", g[["Matricula","aderente","total","pct"]])

def parse_etit_outage_sem_sinal_gpon_from_residencial(planilha_residencial_xlsx, equipe_ids: set[str]) -> Optional[IndicadorResult]:
    """
    ETIT Outage Sem Sinal (GPON) vem da planilha 'Analítico Indicadores Residencial' (aba Analitico).
    Mapeamento adotado:
      - INDICADOR_NOME_ICG == 'ETIT GPON'
      - SINTOMA == 'INTERRUPCAO' (equivalente a Sem Sinal na maioria dos relatórios)
      - % = aderente/total * 100 (INDICADOR=1 é aderente)
    """
    df = pd.read_excel(planilha_residencial_xlsx, sheet_name="Analitico")
    required = {"INDICADOR_NOME_ICG", "INDICADOR", "LOGIN_ACIONAMENTO", "SINTOMA"}
    # Alguns relatórios usam RESPONSAVEL/LOGIN diferente; tentamos achar coluna de matrícula.
    # Se não existir LOGIN_ACIONAMENTO, tenta 'RESPONSAVEL' ou 'LOGIN'.
    if "LOGIN_ACIONAMENTO" not in df.columns:
        for alt in ("RESPONSAVEL","LOGIN"):
            if alt in df.columns:
                df = df.rename(columns={alt:"LOGIN_ACIONAMENTO"})
                break

    required = {"INDICADOR_NOME_ICG", "INDICADOR", "LOGIN_ACIONAMENTO", "SINTOMA"}
    if not required.issubset(set(df.columns)):
        return None

    nome = _safe_upper(df["INDICADOR_NOME_ICG"])
    sint = _safe_upper(df["SINTOMA"])

    sub = df[(nome == RES_ETIT_GPON_INDICADOR_NOME) & (sint == RES_SINTOMA_SEM_SINAL)].copy()
    if sub.empty:
        return None

    sub = sub[sub["LOGIN_ACIONAMENTO"].isin(equipe_ids)]
    if sub.empty:
        return None

    g = sub.groupby("LOGIN_ACIONAMENTO").agg(total=("INDICADOR","size"), aderente=("INDICADOR","sum")).reset_index()
    g["pct"] = _calc_pct(g["aderente"], g["total"])
    g = g.rename(columns={"LOGIN_ACIONAMENTO":"Matricula"})
    return IndicadorResult("ETIT Outage Sem Sinal (GPON)", g[["Matricula","aderente","total","pct"]])

def parse_log_reprog_gpon_from_residencial(planilha_residencial_xlsx, equipe_ids: set[str]) -> Optional[IndicadorResult]:
    """
    Log Outage Reprog. GPON (meta 10%, menor é melhor).
    Na planilha residencial, normalmente:
      - INDICADOR_NOME_ICG == 'LOG REPROGRAMAÇÃO GPON'
      - SINTOMA == 'INTERRUPCAO' (sem sinal)
    Para respeitar a direção (↓), aqui calculamos:
      - pct = (% de NÃO ADERENTE) = (total - aderente)/total * 100
    Assim, quanto menor o pct, melhor.
    """
    df = pd.read_excel(planilha_residencial_xlsx, sheet_name="Analitico")
    if "LOGIN_ACIONAMENTO" not in df.columns:
        for alt in ("RESPONSAVEL","LOGIN"):
            if alt in df.columns:
                df = df.rename(columns={alt:"LOGIN_ACIONAMENTO"})
                break

    required = {"INDICADOR_NOME_ICG", "INDICADOR", "LOGIN_ACIONAMENTO", "SINTOMA"}
    if not required.issubset(set(df.columns)):
        return None

    nome = _safe_upper(df["INDICADOR_NOME_ICG"])
    sint = _safe_upper(df["SINTOMA"])

    sub = df[(nome == RES_LOG_REPROG_GPON_INDICADOR_NOME) & (sint == RES_SINTOMA_SEM_SINAL)].copy()
    if sub.empty:
        return None

    sub = sub[sub["LOGIN_ACIONAMENTO"].isin(equipe_ids)]
    if sub.empty:
        return None

    g = sub.groupby("LOGIN_ACIONAMENTO").agg(total=("INDICADOR","size"), aderente=("INDICADOR","sum")).reset_index()
    g["nao_aderente"] = g["total"] - g["aderente"]
    g["pct"] = _calc_pct(g["nao_aderente"], g["total"])
    g = g.rename(columns={"LOGIN_ACIONAMENTO":"Matricula"})
    # manter colunas padrão: aderente, total, pct (aqui 'aderente' vira "não reprogramou" se sua regra for essa)
    return IndicadorResult("Log Outage Reprog. GPON", g[["Matricula","aderente","total","pct"]])
