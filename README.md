# 🎟️ Distribuidor de Senhas — Streamlit

App em Streamlit que lê uma planilha do Google Sheets e distribui **senhas sequenciais por área**,
gravando os dados e gerando um **PDF** pronto para impressão.

## ✅ Estrutura da Planilha

- Aba **`Nomes`** (editável): deve conter ao menos as colunas:
  - `Área` — nome exibido no app
  - `Aba` *(opcional)* — nome da aba de destino; se ausente, usa o próprio texto de `Área`
  - `Ativa` — *Sim/Não* (ou True/False, 1/0)

- Para **cada área ativa**, o app grava **nessa aba** (criando se não existir) o seguinte cabeçalho:
  ```
  Senha | Nome | Telefone | Bairro | Data e Hora de Registro | Data e Hora de Atendimento
  ```

A *Senha* é sequencial por planilha (linha - 1, considerando a linha 1 como cabeçalho).

## 🔐 Segredos (Streamlit Cloud ou local)

No arquivo `.streamlit/secrets.toml` defina:

```toml
SPREADSHEET_ID = "SUA_PLANILHA_ID_AQUI"

# Preferível em produção: conta de serviço
GOOGLE_SERVICE_ACCOUNT_JSON = """
{...json da conta de serviço...}
"""

# Alternativa: OAuth de usuário (não recomendado para multiusuário)
# GOOGLE_CLIENT_SECRET = """
# {...}
# """
```

> Dica: compartilhe a planilha com o e-mail da conta de serviço com permissão de **Editor**.

## ▶️ Rodando

- Local: `pip install -r requirements.txt` e depois `streamlit run streamlit_app_senhas.py`
- Cloud: suba estes arquivos e configure `secrets.toml` conforme acima.

## 🧱 Base / Inspiração

- Padrão de autenticação e escrita no Sheets e técnica para extrair a linha gravada via `updatedRange` foram inspirados dos utilitários existentes (ver `utils.py` e `streamlit_app.py`).

## 🖼️ Logotipo do PDF

Para personalizar o cabeçalho do ticket, coloque um arquivo `logo.png` dentro da pasta `assets/`. O arquivo é lido em tempo de execução e **não precisa (nem deve) ser versionado**: ele já está listado no `.gitignore`, então faça o upload manual no ambiente de execução.

Se preferir outro caminho, defina a variável de ambiente `PDF_LOGO_PATH` apontando para o arquivo desejado.
