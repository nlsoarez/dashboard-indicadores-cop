# Dashboard de Indicadores (Upload de Planilhas)

Este projeto cria um dashboard simples (Streamlit) para:
- Fazer upload das planilhas
- Extrair os indicadores por **matrícula** (somente sua equipe)
- Comparar automaticamente com as **metas**
- Exibir tabela, alertas e ranking

## Indicadores implementados (neste repo)
- **Chat TOA** (best-effort na planilha TOA — usa padrões de texto; se o mês não vier com o indicador, ele não aparece)
- **ETIT Outage Sem Sinal (GPON)** (planilha Residencial)
- **Log Outage Reprog. GPON** (planilha Residencial; meta 10%, menor é melhor)

> Se você quiser, dá pra plugar os demais indicadores no mesmo padrão criando novos parsers em `src/parsers.py`.

---

## Como rodar local

### 1) Criar venv e instalar
```bash
python -m venv .venv
# Windows:
.venv\Scripts\activate
# Linux/Mac:
source .venv/bin/activate

pip install -r requirements.txt
```

### 2) Rodar
```bash
streamlit run app.py
```

### 3) Upload das planilhas
Na tela, faça upload de:
- **Analitico Indicadores TOA** (aba `TOA`)
- **Analítico Indicadores Residencial** (aba `Analitico`)

Clique em **Processar Dados**.

---

## Estrutura do projeto
- `app.py` → dashboard (UI)
- `src/config.py` → equipe, metas e padrões
- `src/parsers.py` → regras de extração dos indicadores
- `data/` → pasta vazia (caso queira colocar exemplos)

---

## Ajustes rápidos (se mudar o nome do indicador)
Em `src/config.py`, altere:
- `CHAT_TOA_NAME_PATTERNS`
- `RES_ETIT_GPON_INDICADOR_NOME`
- `RES_LOG_REPROG_GPON_INDICADOR_NOME`
- `RES_SINTOMA_SEM_SINAL`
