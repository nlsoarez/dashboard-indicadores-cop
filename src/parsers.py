import pandas as pd
from src.config import CHAT_TOA_NAME_PATTERNS


def _norm(s: str) -> str:
    return str(s).strip().upper()


def find_col(df: pd.DataFrame, candidates):
    """
    Encontra coluna por "nome parecido", ignorando case e espaços.
    candidates: lista de strings que podem existir no df.
    """
    cols = { _norm(c): c for c in df.columns }
    for cand in candidates:
        cand_norm = _norm(cand)
        if cand_norm in cols:
            return cols[cand_norm]

    # fallback: contém substring
    for real in df.columns:
        r = _norm(real)
        for cand in candidates:
            if _norm(cand) in r:
                return real

    return None


def status_por_meta(valor_pct, meta, direcao):
    if valor_pct is None or (isinstance(valor_pct, float) and pd.isna(valor_pct)):
        return "—"
    if direcao == "up":
        return "✅ Dentro" if valor_pct >= meta else "❌ Fora"
    return "✅ Dentro" if valor_pct <= meta else "❌ Fora"


def parse_chat_toa_from_df(df_toa: pd.DataFrame, equipe_ids: set) -> pd.DataFrame:
    """
    Esperado (TOA):
    - Uma coluna com matrícula/login (ex.: LOGIN)
    - Uma coluna com nome do indicador (ex.: INDICADOR_NOME)
    - Uma coluna com aderência 0/1 (ex.: INDICADOR)
    """
    login_col = find_col(df_toa, ["LOGIN", "MATRICULA", "MATRÍCULA", "USUARIO", "USUÁRIO"])
    nome_col = find_col(df_toa, ["INDICADOR_NOME", "INDICADOR", "NOME_INDICADOR", "INDICADOR NOME"])
    val_col = find_col(df_toa, ["INDICADOR", "VALOR", "RESULTADO", "ADERENCIA", "ADERÊNCIA"])

    if not login_col or not nome_col or not val_col:
        # não quebra: retorna vazio para aparecer na tabela como NaN
        return pd.DataFrame({"Matricula": list(equipe_ids), "Chat_TOA_pct": [pd.NA] * len(equipe_ids)})

    df = df_toa.copy()
    df[login_col] = df[login_col].astype(str).str.strip()
    df = df[df[login_col].isin(equipe_ids)]

    # filtra registros que "parecem chat"
    nome_norm = df[nome_col].astype(str).apply(_norm)
    mask = nome_norm.apply(lambda x: any(p in x for p in CHAT_TOA_NAME_PATTERNS))
    df = df[mask]

    if df.empty:
        return pd.DataFrame({"Matricula": list(equipe_ids), "Chat_TOA_pct": [pd.NA] * len(equipe_ids)})

    df[val_col] = pd.to_numeric(df[val_col], errors="coerce")

    g = df.groupby(login_col).agg(total=(val_col, "size"), aderente=(val_col, "sum")).reset_index()
    g["Chat_TOA_pct"] = (g["aderente"] / g["total"]) * 100
    g = g.rename(columns={login_col: "Matricula"})
    return g[["Matricula", "Chat_TOA_pct"]]


def parse_gpon_from_residencial_df(df_res: pd.DataFrame, equipe_ids: set) -> pd.DataFrame:
    """
    Esperado (Residencial - Analitico):
    - login/matricula (LOGIN)
    - nome do indicador (INDICADOR_NOME)
    - valor 0/1 (INDICADOR)
    - sintoma (SINTOMA) -> filtrar INTERRUPCAO (Sem sinal)
    Indicadores:
      - ETIT GPON
      - LOG REPROGRAMAÇÃO GPON
    Reprog: meta 10% menor é melhor -> calculamos % NÃO aderente (1 - aderente).
    """
    login_col = find_col(df_res, ["LOGIN", "MATRICULA", "MATRÍCULA", "USUARIO", "USUÁRIO", "LOGIN_ACIONAMENTO"])
    nome_col = find_col(df_res, ["INDICADOR_NOME", "NOME_INDICADOR", "INDICADOR"])
    val_col = find_col(df_res, ["INDICADOR", "VALOR", "RESULTADO", "ADERENCIA", "ADERÊNCIA"])
    sint_col = find_col(df_res, ["SINTOMA", "SINTOMA_OUTAGE", "SINTOMA OUTAGE"])

    if not login_col or not nome_col or not val_col:
        return pd.DataFrame({"Matricula": list(equipe_ids), "ETIT_GPON_pct": [pd.NA]*len(equipe_ids), "REPROG_GPON_pct": [pd.NA]*len(equipe_ids)})

    df = df_res.copy()
    df[login_col] = df[login_col].astype(str).str.strip()
    df = df[df[login_col].isin(equipe_ids)]
    df[val_col] = pd.to_numeric(df[val_col], errors="coerce")

    # Filtrar sintoma interrupção (se existir coluna)
    if sint_col:
        s = df[sint_col].astype(str).apply(_norm)
        df = df[s.str.contains("INTERRUP", na=False)]  # pega INTERRUPCAO / INTERRUPÇÃO

    # helpers
    def pct_aderente(indicador_substring: str):
        sub = df[df[nome_col].astype(str).apply(_norm).str.contains(indicador_substring, na=False)]
        if sub.empty:
            return pd.DataFrame(columns=["Matricula", "pct"])
        g = sub.groupby(login_col).agg(total=(val_col, "size"), aderente=(val_col, "sum")).reset_index()
        g["pct"] = (g["aderente"] / g["total"]) * 100
        g = g.rename(columns={login_col: "Matricula"})
        return g[["Matricula", "pct"]]

    # ETIT GPON
    etit = pct_aderente("ETIT GPON").rename(columns={"pct": "ETIT_GPON_pct"})

    # LOG REPROGRAMAÇÃO GPON
    # Queremos % de reprog (não aderente), então: 100 - %aderente
    rep = pct_aderente("LOG REPROGR").rename(columns={"pct": "aderente_pct"})
    if not rep.empty:
        rep["REPROG_GPON_pct"] = 100 - rep["aderente_pct"]
        rep = rep[["Matricula", "REPROG_GPON_pct"]]
    else:
        rep = pd.DataFrame(columns=["Matricula", "REPROG_GPON_pct"])

    out = pd.DataFrame({"Matricula": list(equipe_ids)})
    out = out.merge(etit, on="Matricula", how="left").merge(rep, on="Matricula", how="left")
    return out
