
# streamlit_app_senhas.py — UI Streamlit para o Distribuidor de Senhas
from __future__ import annotations
from typing import List, Dict
import inspect
import re
import streamlit as st
import os, requests

PRINT_SERVER_URL = st.secrets.get("PRINT_SERVER_URL") or os.getenv("PRINT_SERVER_URL", "")
PRINT_TOKEN      = st.secrets.get("PRINT_TOKEN")      or os.getenv("PRINT_TOKEN", "")


def enviar_para_impressao(pdf_bytes: bytes) -> tuple[bool, str | None]:
    if not (PRINT_SERVER_URL and PRINT_TOKEN):
        return False, "PRINT_SERVER_URL/PRINT_TOKEN não configurados."
    try:
        url = PRINT_SERVER_URL.rstrip("/") + "/print/pdf"
        r = requests.post(
            url,
            headers={"X-Token": PRINT_TOKEN, "Content-Type": "application/pdf"},
            data=pdf_bytes,
            timeout=30,
        )
        if r.ok:
            return True, None
        return False, f"Falha HTTP {r.status_code}: {r.text}"
    except Exception as e:
        return False, str(e)

from event_utils import (
    read_active_areas,
    read_neighborhoods,
    submit_tickets,
    format_phone_number,
    _sheets_service,
    _get_spreadsheet_id,
)

st.set_page_config(page_title="Distribuidor de Senhas — Evento", page_icon="🎟️", layout="centered")
st.title("🎟️ Distribuidor de Senhas — Evento")

st.caption(f"Planilha conectada: `{_get_spreadsheet_id()}` (definida no código)")

# Ajuda rápida
with st.expander("Como funciona?"):
    st.markdown(
        """
        1. A aba **Nomes** da planilha deve listar todas as áreas, com a coluna **Ativa** marcada para as que devem aparecer aqui.
        2. Marque uma ou mais **Áreas** (apenas as ativas são exibidas), preencha **Nome**, **Telefone** e **Bairro**.
        3. Clique em **Gerar senhas e salvar**. O app:
           - grava cada registro na aba correspondente com as colunas `Senha | Nome | Telefone | Bairro | Data e Hora de Registro | Data e Hora de Atendimento` (esta última em branco);
           - cria a **Senha sequencial** da planilha (1, 2, 3, …) para cada área;
           - gera um **PDF** com uma página para cada senha.
        """
    )

# Teste de credenciais e carregamento de áreas
areas_opts: List[Dict] = []
bairros_opts: List[str] = []
try:
    service = _sheets_service()
    sid = _get_spreadsheet_id()
    areas_opts = read_active_areas(service, sid)
    bairros_opts = read_neighborhoods(service, sid)
except Exception as e:
    st.error(f"⚠️ Não foi possível ler a planilha: {e}")

if not areas_opts:
    st.warning("Nenhuma área ativa encontrada na aba 'Nomes'. Verifique a planilha/credenciais.")
else:
    labels = [a["area"] for a in areas_opts]
    areas_sel = st.multiselect(
        "Áreas / Setores",
        options=labels,
        placeholder="Escolha uma ou mais áreas",
    )
    nome_input = st.text_input("Nome", max_chars=80)
    nome = nome_input.strip()
    telefone_input = st.text_input("Telefone", max_chars=30, placeholder="92981231234")
    telefone_feedback = st.empty()
    rede_social_input = st.text_input(
        "Rede social (@...)",
        max_chars=80,
        placeholder="@seudominio",
        help="Opcional. Informe o usuário ou perfil principal.",
    )
    email_input = st.text_input(
        "E-mail",
        max_chars=120,
        placeholder="nome@exemplo.com",
        help="Opcional. Será armazenado apenas na planilha.",
    )
    telefone_ok = True
    telefone_msg = ""
    telefone_preview = ""
    if telefone_input.strip():
        try:
            telefone_preview = format_phone_number(telefone_input)
        except ValueError as exc:
            telefone_ok = False
            telefone_msg = str(exc)
    else:
        telefone_ok = False
        telefone_msg = "Informe o telefone com 11 dígitos (incluindo DDD)."

    if telefone_msg:
        telefone_feedback.caption(f"ℹ️ {telefone_msg}")
    elif telefone_preview:
        telefone_feedback.caption(f"Formato final: {telefone_preview}")
    if bairros_opts:
        bairro = st.selectbox("Bairro", options=[""] + bairros_opts, index=0)
    else:
        st.info(
            "Lista de bairros não encontrada na aba 'Bairro'. Informe manualmente abaixo ou verifique a planilha."
        )
        bairro = st.text_input("Bairro", max_chars=80)

    btn = st.button(
        "✅ Gerar senhas e salvar",
        type="primary",
        disabled=(not areas_sel or not nome or not telefone_ok),
    )

    if btn:
        with st.spinner("Gravando na planilha e gerando PDF..."):
            try:
                submit_kwargs = {
                    "areas": areas_sel,
                    "nome": nome,
                    "telefone": telefone_input,
                    "bairro": bairro,
                }

                try:
                    params = inspect.signature(submit_tickets).parameters
                except (ValueError, TypeError):
                    params = {}

                if "rede_social" in params:
                    submit_kwargs["rede_social"] = rede_social_input
                if "email" in params:
                    submit_kwargs["email"] = email_input

                resultados, pdf_bytes, excedidas = submit_tickets(**submit_kwargs)
                linhas = [
                    f"* Área **{item['area']}** → senha **{item['senha']}** (registro {item['ts_registro']})."
                    for item in resultados
                ]
                st.success("Senhas registradas com sucesso!")
                st.markdown("\n".join(linhas))

                if excedidas:
                    avisos = [
                        (
                            f"Área **{info['area']}** excedeu o limite de {info['limite']} "
                            f"senhas (atual: {info['senha']})."
                        )
                        for info in excedidas
                    ]
                    st.warning(
                        "\n".join(
                            [
                                "⚠️ O PDF não foi gerado porque os limites abaixo foram atingidos:",
                                *avisos,
                            ]
                        )
                    )
                elif pdf_bytes:
                    if len(resultados) == 1:
                        area_nome = resultados[0]["area"]
                        senha_num = resultados[0]["senha"]
                        base_name = f"senha_{area_nome}_{senha_num}"
                    else:
                        base_name = f"senhas_{len(resultados)}_areas"
                    safe_name = re.sub(r"[^A-Za-z0-9_-]+", "_", base_name).strip("_") or "senhas"
                    file_name = f"{safe_name}.pdf"

                    st.download_button(
                        "⬇️ Baixar PDF das senhas",
                        data=pdf_bytes,
                        file_name=file_name,
                        mime="application/pdf",
                    )
                    ok, err = enviar_para_impressao(pdf_bytes)
                    if ok:
                        st.success("🖨️ Enviado automaticamente para impressão.")
                    else:
                        st.warning(f"Não foi possível imprimir automaticamente: {err}")
            except ValueError as e:
                st.error(str(e))
            except Exception as e:
                st.error(f"Falha ao gerar senha: {e}")
