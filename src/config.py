from __future__ import annotations

# ---------------------------
# Base fixa: equipe e metas
# ---------------------------

EQUIPE = [
    ("N6088107","LEANDRO GONÇALVES DE CARVALHO","EMPRESARIAL"),
    ("N5619600","BRUNO COSTA BUCARD","EMPRESARIAL"),
    ("N0189105","IGOR MARCELINO DE MARINS","EMPRESARIAL"),
    ("N5737414","SANDRO DA SILVA CARVALHO","EMPRESARIAL"),
    ("N5713690","GABRIELA TAVARES DA SILVA","EMPRESARIAL"),
    ("N5802257","MAGNO FERRAREZ DE MORAIS","EMPRESARIAL"),
    ("F201714","FERNANDA MESQUITA DE FREITAS","EMPRESARIAL"),
    ("N6173055","JEFFERSON LUIS GONÇALVES COITINHO","EMPRESARIAL"),
    ("N0125317","ROBERTO SILVA DO NASCIMENTO","EMPRESARIAL"),
    ("F218860","ALDENES MARQUES IDALINO DA SILVA","EMPRESARIAL"),
    ("N5819183","RODRIGO PIRES BERNARDINO","EMPRESARIAL"),
    ("N5926003","SUELLEN HERNANDEZ DA SILVA","EMPRESARIAL"),
    ("N5932064","MONICA DA SILVA RODRIGUES","EMPRESARIAL"),
    ("N0238475","MARLEY MARQUES RIBEIRO","RESIDENCIAL"),
    ("N5923221","KELLY PINHEIRO LIRA","RESIDENCIAL"),
    ("N5772086","THIAGO PEREIRA DA SILVA","RESIDENCIAL"),
    ("N0239871","LEONARDO FERREIRA LIMA DE ALMEIDA","RESIDENCIAL"),
    ("N5577565","MARISTELLA MARCIA DOS SANTOS","RESIDENCIAL"),
    ("N5972428","CRISTIANE HERMOGENES DA SILVA","RESIDENCIAL"),
    ("N4014011","ALAN MARINHO DIAS","RESIDENCIAL"),
    ("F106664","RAISSA LIMA DE OLIVEIRA","RESIDENCIAL"),
]

# Metas (em %)
METAS = {
    "DPA - Ocupação": {"meta": 90, "direcao": "up"},
    "Chat TOA": {"meta": 75, "direcao": "up"},
    "ETIT por Evento RAL": {"meta": 90, "direcao": "up"},
    "ETIT por Evento REC": {"meta": 90, "direcao": "up"},
    "ETIT Outage Sem Sinal (GPON)": {"meta": 90, "direcao": "up"},
    "Log Outage Reprog. GPON": {"meta": 10, "direcao": "down"},
}

# Como identificar os 3 indicadores dentro das duas planilhas
# Obs.: Chat TOA às vezes vem com nomes diferentes. Coloquei padrões (contains) para funcionar mesmo mudando texto.
CHAT_TOA_NAME_PATTERNS = [
    "CHAT", "INTERAÇÃO", "INTERACAO", "10 MIN", "10MIN", "TOA CHAT"
]

# Na planilha Residencial (aba Analitico)
RES_ETIT_GPON_INDICADOR_NOME = "ETIT GPON"
RES_LOG_REPROG_GPON_INDICADOR_NOME = "LOG REPROGRAMAÇÃO GPON"
RES_SINTOMA_SEM_SINAL = "INTERRUPCAO"  # costuma ser o equivalente ao "Sem Sinal"
