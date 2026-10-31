from __future__ import annotations

import hmac
import os
import sys
from pathlib import Path

import streamlit as st

APP_DIR = Path(__file__).resolve().parents[1]
if str(APP_DIR) not in sys.path:
    sys.path.insert(0, str(APP_DIR))

from Home import (  # noqa: E402
    apply_store_source,
    current_manifest,
    load_cache,
    load_settings,
    log_sync,
    manifest_signature,
    online_database_url,
    save_settings,
    sync_to_postgres,
    sync_settings_to_postgres,
)


st.set_page_config(page_title="Configuracoes | Desmobilizacao", layout="wide")


def configured_password() -> str:
    env_password = os.getenv("SETTINGS_PAGE_PASSWORD", "").strip()
    if env_password:
        return env_password
    try:
        return str(st.secrets.get("settings_password", "")).strip()
    except Exception:
        return ""


def require_settings_password() -> None:
    if st.session_state.get("settings_page_authorized"):
        return
    st.title("Configuracoes")
    st.caption("Esta pagina e protegida por senha.")
    with st.form("settings_page_login"):
        password = st.text_input("Senha das configuracoes", type="password")
        submitted = st.form_submit_button("Acessar", type="primary", width="stretch")
    if submitted:
        expected = configured_password()
        if not expected:
            st.error("Defina settings_password nos secrets do Streamlit ou SETTINGS_PAGE_PASSWORD no ambiente.")
        elif hmac.compare_digest(password, expected):
            st.session_state["settings_page_authorized"] = True
            st.rerun()
        else:
            st.error("Senha incorreta.")
    st.stop()


require_settings_password()
st.title("Configuracoes do relatorio")
st.caption("As alteracoes ficam salvas no banco SQLite local.")
settings = load_settings()
remote_url = online_database_url()

with st.form("report_settings"):
    source_index = 1 if settings["store_count_source"] == "custo" else 0
    store_source = st.selectbox(
        "Origem da vinculacao das 298 lojas",
        ["transp", "custo"],
        index=source_index,
        format_func=lambda value: "Aba Transp" if value == "transp" else "Aba Custo Frete",
    )
    codes_80 = st.text_input("BOA para QEB (Reason Code 80)", value=settings["codes_boa_to_qeb"])
    codes_82 = st.text_input("QEB para BOA (Reason Code 82)", value=settings["codes_qeb_to_boa"])
    codes_t1 = st.text_input("QEB para SLD (Reason Code T1)", value=settings["codes_qeb_to_sld"])
    codes_av = st.text_input("Retorno sem debito (Reason Code AV)", value=settings["codes_av_return"])
    submitted = st.form_submit_button("Salvar configuracoes", type="primary", width="stretch")

if submitted:
    updated_settings = {
        "store_count_source": store_source,
        "codes_boa_to_qeb": codes_80,
        "codes_qeb_to_boa": codes_82,
        "codes_qeb_to_sld": codes_t1,
        "codes_av_return": codes_av,
    }
    save_settings(updated_settings)
    if remote_url:
        try:
            sync_settings_to_postgres(remote_url, updated_settings)
            st.success("Configuracoes salvas no SQLite e no PostgreSQL online.")
        except Exception as error:
            st.warning(f"Configuracoes salvas no SQLite; falha ao salvar online: {error}")
    else:
        st.success("Configuracoes salvas no SQLite.")
    st.rerun()

st.divider()
st.subheader("Banco de dados")
st.caption("SQLite local: configuracoes e dados agregados ficam salvos junto ao projeto.")
if remote_url:
    st.success("PostgreSQL online configurado.")
    manifest = current_manifest()
    signature = manifest_signature(manifest) if manifest else ""
    payload = load_cache(signature) if signature else None
    if payload is None:
        st.info("Acesse a pagina principal e atualize os dados antes de sincronizar.")
    elif st.button("Enviar dados agregados ao PostgreSQL", type="primary", width="stretch"):
        try:
            mapped_payload = apply_store_source(payload, settings["store_count_source"])
            if mapped_payload["validation_issues"]:
                st.error("Corrija as pendencias das bases antes de sincronizar: " + "; ".join(mapped_payload["validation_issues"]))
            else:
                sync_to_postgres(remote_url, mapped_payload, signature)
                log_sync(signature, "enviado")
                st.success("Dados agregados enviados ao PostgreSQL online.")
        except Exception as error:
            log_sync(signature, f"erro: {error}")
            st.error(f"Falha na sincronizacao: {error}")
else:
    st.info("Sincronizacao online opcional. Configure database_url nos secrets para habilitar o PostgreSQL.")

if st.button("Bloquear configuracoes"):
    st.session_state["settings_page_authorized"] = False
    st.rerun()
