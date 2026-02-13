import pandas as pd

# =====================================================
# EQUIPE - Matrículas e Nomes
# =====================================================
EQUIPE = [
    # EMPRESARIAL
    ("N6088107", "LEANDRO GONÇALVES DE CARVALHO", "EMPRESARIAL"),
    ("N5619600", "BRUNO COSTA BUCARD", "EMPRESARIAL"),
    ("N0189105", "IGOR MARCELINO DE MARINS", "EMPRESARIAL"),
    ("N5737414", "SANDRO DA SILVA CARVALHO", "EMPRESARIAL"),
    ("N5713690", "GABRIELA TAVARES DA SILVA", "EMPRESARIAL"),
    ("N5802257", "MAGNO FERRAREZ DE MORAIS", "EMPRESARIAL"),
    ("F201714", "FERNANDA MESQUITA DE FREITAS", "EMPRESARIAL"),
    ("N6173055", "JEFFERSON LUIS GONÇALVES COITINHO", "EMPRESARIAL"),
    ("N0125317", "ROBERTO SILVA DO NASCIMENTO", "EMPRESARIAL"),
    ("F218860", "ALDENES MARQUES IDALINO DA SILVA", "EMPRESARIAL"),
    ("N5819183", "RODRIGO PIRES BERNARDINO", "EMPRESARIAL"),
    ("N5926003", "SUELLEN HERNANDEZ DA SILVA", "EMPRESARIAL"),
    ("N5932064", "MONICA DA SILVA RODRIGUES", "EMPRESARIAL"),
    # RESIDENCIAL
    ("N0238475", "MARLEY MARQUES RIBEIRO", "RESIDENCIAL"),
    ("N5923221", "KELLY PINHEIRO LIRA", "RESIDENCIAL"),
    ("N5772086", "THIAGO PEREIRA DA SILVA", "RESIDENCIAL"),
    ("N0239871", "LEONARDO FERREIRA LIMA DE ALMEIDA", "RESIDENCIAL"),
    ("N5577565", "MARISTELLA MARCIA DOS SANTOS", "RESIDENCIAL"),
    ("N5972428", "CRISTIANE HERMOGENES DA SILVA", "RESIDENCIAL"),
    ("N4014011", "ALAN MARINHO DIAS", "RESIDENCIAL"),
    ("F106664", "RAISSA LIMA DE OLIVEIRA", "RESIDENCIAL"),
]

BASE_EQUIPE = pd.DataFrame(EQUIPE, columns=["Matricula", "Nome", "Setor"])
EQUIPE_IDS = set(BASE_EQUIPE["Matricula"].tolist())

# =====================================================
# COLUNAS ESPERADAS NA PLANILHA DE PRODUTIVIDADE
# =====================================================
COL_LOGIN = "USUARIO_LOGIN"
COL_NOME = "USUARIO_NOME"
COL_BASE = "USUARIO_BASE"
COL_COORD = "USUARIO_COORD"
COL_CARGO = "USUARIO_CARGO"
COL_PERIODO = "USUARIO_PERIODO"
COL_DATA = "DATA"
COL_MES = "MESNOME"
COL_ANOMES = "ANOMES"

# Volumes
VOL_COLS = {
    "VOL_AB_NM": "Abertura New Monitor",
    "VOL_FE_NM": "Fechamento New Monitor",
    "VOL_FE_NM_MANOBRA": "Fechamento NM Manobra",
    "VOL_AB_SGO": "Abertura SGO",
    "VOL_TRAT_SGO": "Tratamento SGO",
    "VOL_AC_SGO": "Aceite SGO",
    "VOL_FE_SGO": "Fechamento SGO",
    "VOL_AB_OSS": "Abertura Remedy",
    "VOL_FE_OSS": "Fechamento OSS",
    "VOL_AC_OSS": "Aceite OSS",
    "VOL_RAL": "Tratativa RAL",
    "VOL_REC": "Tratativa REC",
    "VOL_AB_RAL": "Abertura RAL",
    "VOL_REMEDY_MOVEL": "Remedy Móvel",
    "VOL_TOA_PRIM_INT": "Primeira Interação TOA",
    "VOL_TOA_FORM": "Fechamento Tarefa TOA",
    "VOL_TELEFONIA_RECEBIDO": "Telefonia Recebido",
    "VOL_TELEFONIA_ATENDIDO": "Telefonia Atendido",
    "VOL_TELEFONIA_REALIZADO": "Ligações Realizadas",
}

# Volume columns grouped by sector
VOL_COLS_RESIDENCIAL = {
    "VOL_AB_NM": "Ab. New Monitor",
    "VOL_FE_NM": "Fech. New Monitor",
    "VOL_AB_SGO": "Ab. SGO",
    "VOL_FE_SGO": "Fech. SGO",
}

VOL_COLS_EMPRESARIAL = {
    "VOL_RAL": "Trat. RAL",
    "VOL_REC": "Trat. REC",
}

VOL_COLS_AMBOS = {
    "VOL_AB_OSS": "Ab. Remedy",
    "VOL_TELEFONIA_REALIZADO": "Ligações Realiz.",
    "VOL_TOA_PRIM_INT": "1ª Interação TOA",
    "VOL_TOA_FORM": "Fech. Tarefa TOA",
}

COL_VOL_TOTAL = "VOL_TOTAL"
COL_VOL_MEDIA = "VOL_TOTAL_MEDIA_POR_DIA"

# DPA (Ocupação)
COL_DPA_USO = "DPA_TEMPO_USO_SEC"
COL_DPA_JORNADA = "DPA_HORARIO_JORNADA_SEC"
COL_DPA_RESULTADO = "DPA_RESULTADO"

# Header row na planilha (0-indexed)
HEADER_ROW = 10

# Nome da aba (fallback: primeira aba)
SHEET_NAME_CANDIDATES = [
    "Analítico Produtividade 2026",
    "Analítico Produtividade",
    "Produtividade",
]

# =====================================================
# CORES DO DASHBOARD
# =====================================================
COR_PRIMARIA = "#1B4F72"
COR_SUCESSO = "#27AE60"
COR_ALERTA = "#F39C12"
COR_PERIGO = "#E74C3C"
COR_INFO = "#2980B9"
