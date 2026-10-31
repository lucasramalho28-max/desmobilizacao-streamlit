# Relatorio de Desmobilizacao

Aplicacao Streamlit independente para acompanhar a desmobilizacao dos CDs e das 298 lojas.

## Rodar localmente

1. Copie `.streamlit/secrets.toml.example` para `.streamlit/secrets.toml`.
2. Configure `app_password` (relatorio) e `settings_password` (configuracoes).
3. Instale e inicie:

```powershell
py -3 -m pip install -r requirements.txt
py -3 -m streamlit run Home.py
```

A pagina inicial e o relatorio. Configure senhas diferentes para o relatorio e para Configuracoes no arquivo de secrets. Os secrets reais nao ficam no Git.

Se `1.09/`, `9.09/` e `Lista final_298 lojas 1.xlsx` estiverem ao lado do app, o botao Atualizar dados os reconhece automaticamente. No Streamlit Cloud, envie um ZIP unico com essas duas pastas e a planilha; os arquivos grandes de origem nao estao neste repositorio.

## Deploy no Streamlit Community Cloud

Crie o app apontando para este repositorio, branch `main`, arquivo `streamlit_app.py` (entrypoint na raiz; ele abre a pagina principal definida em `Home.py`). Em Advanced settings > Secrets, informe:

```toml
app_password = "SENHA_DO_RELATORIO"
settings_password = "SENHA_DAS_CONFIGURACOES"
database_url = "postgresql://USUARIO:SENHA@HOST:5432/BANCO"
```

`database_url` e opcional para abrir o relatorio; configure PostgreSQL se precisar persistir os agregados e as configuracoes entre reinicios do app. A pagina protegida de Configuracoes permite sincronizar os agregados validados; depois disso, a aplicacao consegue recuperar o ultimo snapshot online.

Os dados de origem, o banco SQLite local e secrets estao excluidos do Git. Consulte a documentacao oficial do Streamlit para [deploy e secrets](https://docs.streamlit.io/deploy/streamlit-community-cloud/deploy-your-app/deploy).
