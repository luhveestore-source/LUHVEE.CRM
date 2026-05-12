import streamlit as st
import pandas as pd

st.set_page_config(page_title="LuhVee CRM PRO 💖", layout="wide")

st.title("💖 LuhVee CRM PRO")

# =========================
# FUNÇÃO LIMPEZA LEADS
# =========================

def limpar_leads(df):

    # renomear colunas manualmente
    novas_colunas = [
        "nome",
        "extra1",
        "tipo_rua",
        "rua",
        "numero",
        "complemento",
        "bairro",
        "cidade",
        "estado",
        "codigo_ibge",
        "cep",
        "telefone",
        "extra2",
        "email",
        "site"
    ]

    df.columns = novas_colunas

    # remover colunas inúteis
    df = df[[
        "nome",
        "telefone",
        "email",
        "cidade",
        "estado",
        "bairro",
        "cep",
        "site"
    ]]

    # limpar vazios
    df = df.fillna("")

    return df

# =========================
# IMPORTAÇÃO EXCEL
# =========================

st.subheader("📂 Importar Leads")

upload = st.file_uploader(
    "Envie Excel ou CSV",
    type=["xlsx", "csv"]
)

if upload is not None:

    try:

        # detectar formato
        if upload.name.endswith("xlsx"):
            df = pd.read_excel(upload)
        else:
            df = pd.read_csv(
                upload,
                encoding="latin1",
                sep=None,
                engine="python"
            )

        # limpar leads
        leads = limpar_leads(df)

        st.success("💖 Leads carregados com sucesso!")

        # métricas
        col1, col2, col3 = st.columns(3)

        with col1:
            st.metric("Total Leads", len(leads))

        with col2:
            st.metric(
                "Com Email",
                leads["email"].astype(bool).sum()
            )

        with col3:
            st.metric(
                "Com Telefone",
                leads["telefone"].astype(bool).sum()
            )

        # filtros
        st.subheader("🔎 Filtrar Leads")

        cidade = st.selectbox(
            "Cidade",
            ["Todas"] + sorted(leads["cidade"].unique().tolist())
        )

        if cidade != "Todas":
            leads = leads[leads["cidade"] == cidade]

        # tabela
        st.dataframe(leads, use_container_width=True)

        # download CSV limpo
        csv = leads.to_csv(index=False).encode("utf-8")

        st.download_button(
            "⬇️ Baixar Leads Organizados",
            csv,
            "leads_luhvee.csv",
            "text/csv"
        )

        # mensagens automáticas
        st.subheader("💬 Mensagem Automática")

        mensagem = """💖 Olá! Tudo bem?

Vi que você gosta de novidades 👠✨

A LuhVee Stores está com modelos lindos disponíveis hoje 💕

Quer receber nosso catálogo?"""

        st.text_area(
            "Mensagem pronta",
            mensagem,
            height=180
        )

    except Exception as e:
        st.error("Erro ao importar leads")
        st.write(e)