# 📊 Dashboard de Produtividade — COP Rede

Dashboard Streamlit para análise de produtividade da equipe COP Rede.
Basta fazer upload da planilha **Produtividade COP Rede - Analítico** e o sistema filtra automaticamente os dados dos seus analistas.

## Funcionalidades

- **Upload único** — suba a planilha e os dados são processados automaticamente
- **Filtros** — por mês, setor (Empresarial/Residencial) e analista individual
- **KPIs** — Volume Total, Analistas Ativos, Média por Analista, DPA (Ocupação)
- **Rankings** — Volume Total, Média Diária, DPA
- **Evolução diária** — gráficos de volume e produtividade ao longo do tempo
- **Composição de volume** — breakdown por tipo de atividade (NM, SGO, OSS, RAL, TOA, Telefonia etc.)
- **Visão individual** — selecione um analista para ver seus dados em detalhe
- **Export CSV** — baixe os dados filtrados

## Equipe monitorada

**Empresarial (13):** Leandro, Bruno, Igor, Sandro, Gabriela, Magno, Fernanda, Jefferson, Roberto, Aldenes, Rodrigo, Suellen, Monica

**Residencial (8):** Marley, Kelly, Thiago, Leonardo, Maristella, Cristiane, Alan, Raissa

## Como rodar

```bash
# Criar venv
python -m venv .venv
source .venv/bin/activate  # Linux/Mac
# .venv\Scripts\activate   # Windows

# Instalar dependências
pip install -r requirements.txt

# Rodar
streamlit run app.py
```

## Estrutura

```
├── app.py                 # Dashboard principal
├── src/
│   ├── config.py          # Equipe, colunas, configurações
│   └── processors.py      # Lógica de processamento dos dados
├── requirements.txt
└── README.md
```

## Ajustes

- Para alterar a equipe, edite `EQUIPE` em `src/config.py`
- Para alterar a linha do header da planilha, edite `HEADER_ROW` em `src/config.py`
- Os nomes das abas aceitas estão em `SHEET_NAME_CANDIDATES`
