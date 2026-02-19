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

# Líderes (supervisores)
LIDERES_IDS = {"N0238475", "N5923221", "N6088107", "N5619600"}  # Marley, Kelly, Leandro, Bruno

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
# ETIT POR EVENTO — Colunas da planilha Analítico Empresarial
# =====================================================
ETIT_INDICADOR_FILTRO = "ETIT POR EVENTO"
ETIT_COL_INDICADOR = "INDICADOR_NOME"
ETIT_COL_LOGIN = "LOGIN_ACIONAMENTO"
ETIT_COL_DEMANDA = "DEMANDA"
ETIT_COL_NOTA = "NOTA"
ETIT_COL_VOLUME = "VOLUME"
ETIT_COL_INDICADOR_VAL = "INDICADOR"
ETIT_COL_STATUS = "INDICADOR_STATUS"
ETIT_COL_TIPO = "TIPO"
ETIT_COL_AREA = "AREA_ENVOLVIDA"
ETIT_COL_CAUSA = "CAUSA"
ETIT_COL_REGIONAL = "IN_REGIONAL"
ETIT_COL_GRUPO = "IN_GRUPO"
ETIT_COL_CIDADE = "IN_CIDADE_UF"
ETIT_COL_UF = "IN_UF"
ETIT_COL_TOA = "ENVIADO_TOA"
ETIT_COL_DT_INICIO = "DT_INICIO"
ETIT_COL_DT_FIM = "DT_FIM"
ETIT_COL_DT_EMISSAO = "DT_EMISSAO"
ETIT_COL_DT_ACIONAMENTO = "DT_ACIONAMENTO"
ETIT_COL_TURNO = "TURNO"
ETIT_COL_TMA = "TMA"
ETIT_COL_TMR = "TMR"
ETIT_COL_ANOMES = "ANOMES"

ETIT_SHEET_CANDIDATES = ["Empresarial", "ETIT", "Analítico"]

# =====================================================
# INDICADORES RESIDENCIAL — Planilha Analítico Indicadores
# =====================================================
RES_IND_ETIT_FIBRA_HFC        = "ETIT FIBRA HFC"
RES_IND_ETIT_GPON              = "ETIT GPON"
RES_IND_REPROG_GPON            = "REPROGRAMAÇÃO GPON"
RES_IND_ASSERT_FIBRA_HFC       = "ASSERTIVIDADE ACIONAMENTO FIBRA HFC"
RES_IND_ASSERT_GPON            = "ASSERTIVIDADE ACIONAMENTO GPON"

RES_INDICADORES_FILTRO = [
    RES_IND_ETIT_FIBRA_HFC,
    RES_IND_ETIT_GPON,
    RES_IND_REPROG_GPON,
    RES_IND_ASSERT_FIBRA_HFC,
    RES_IND_ASSERT_GPON,
]

RES_IND_LABELS = {
    RES_IND_ETIT_FIBRA_HFC:   "ETIT Fibra HFC",
    RES_IND_ETIT_GPON:         "ETIT GPON",
    RES_IND_REPROG_GPON:       "Reprog. GPON",
    RES_IND_ASSERT_FIBRA_HFC:  "Assert. Fibra HFC",
    RES_IND_ASSERT_GPON:       "Assert. GPON",
}

RES_IND_COLORS = {
    RES_IND_ETIT_FIBRA_HFC:   "#E67E22",
    RES_IND_ETIT_GPON:         "#8E44AD",
    RES_IND_REPROG_GPON:       "#2980B9",
    RES_IND_ASSERT_FIBRA_HFC:  "#27AE60",
    RES_IND_ASSERT_GPON:       "#16A085",
}

RES_IND_INVERTIDOS = {RES_IND_REPROG_GPON}

RES_COL_INDICADOR_NOME  = "INDICADOR_NOME_ICG"
RES_COL_ID_MOSTRA       = "ID_MOSTRA"
RES_COL_VOLUME          = "VOLUME"
RES_COL_INDICADOR_VAL   = "INDICADOR"
RES_COL_STATUS          = "INDICADOR_STATUS"
RES_COL_REGIONAL        = "IN_REGIONAL"
RES_COL_GRUPO           = "IN_GRUPO"
RES_COL_CIDADE          = "IN_CIDADE_UF"
RES_COL_UF              = "IN_UF"
RES_COL_TECNOLOGIA      = "TECNOLOGIA"
RES_COL_SERVICO         = "SERVICO"
RES_COL_NATUREZA        = "NATUREZA"
RES_COL_SINTOMA         = "SINTOMA"
RES_COL_FERRAMENTA      = "FERRAMENTA_ABERTURA"
RES_COL_FECHAMENTO      = "FECHAMENTO"
RES_COL_SOLUCAO         = "SOLUCAO"
RES_COL_IMPACTO         = "IMPACTO"
RES_COL_ENVIADO_TOA     = "ENVIADO_TOA"
RES_COL_DT_INICIO       = "DT_INICIO"
RES_COL_DT_FIM          = "DT_FIM"
RES_COL_TMA             = "TMA"
RES_COL_TMR             = "TMR"
RES_COL_ANOMES          = "ANOMES"

RES_SHEET_CANDIDATES = ["Analitico", "Analítico", "Residencial", "Sheet1"]

# =====================================================
# OCUPAÇÃO DPA — Planilha Ocupação_DPA_2026
# =====================================================

# Nomes das abas na planilha de Ocupação DPA
DPA_SHEET_ANALISTAS   = "Analistas"
DPA_SHEET_CONSOLIDADO = "Consolidado"

# Lista de meses em português (para detectar o mês mais recente)
DPA_MESES_PT = [
    "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
    "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro",
]

# Thresholds de semáforo para DPA%
DPA_THRESHOLD_OK      = 90.0   # >= verde
DPA_THRESHOLD_ALERTA  = 80.0   # >= amarelo, < verde
# abaixo de DPA_THRESHOLD_ALERTA = vermelho

# Abas possíveis (candidatos, caso os nomes mudem)
DPA_SHEET_CANDIDATES = ["Analistas", "Analisats", "Analistas DPA"]

# =====================================================
# CORES DO DASHBOARD
# =====================================================
COR_PRIMARIA = "#1B4F72"
COR_SUCESSO = "#27AE60"
COR_ALERTA = "#F39C12"
COR_PERIGO = "#E74C3C"
COR_INFO = "#2980B9"
