import streamlit as st
import pandas as pd

# banco simples
if "leads" not in st.session_state:
    st.session_state.leads = []

st.title("💖 LuhVee CRM - Leads")

st.subheader("Adicionar Lead")

nome = st.text_input("Nome")
whatsapp = st.text_input("WhatsApp")
email = st.text_input("Email")
interesse = st.selectbox("Interesse", ["Sapatos", "Botas", "Scarpin", "Sandálias"])

if st.button("Salvar Lead 💖"):
    st.session_state.leads.append({
        "Nome": nome,
        "WhatsApp": whatsapp,
        "Email": email,
        "Interesse": interesse
    })
    st.success("Lead salvo com sucesso!")

st.subheader("Lista de Leads")

df = pd.DataFrame(st.session_state.leads)
st.dataframe(df)

st.subheader("Mensagem WhatsApp automática")

mensagem = """
Olá 💖

Aqui é da LuhVee Stores Shoes 👠✨
Temos novidades incríveis pra você!

Me chama aqui 👉 https://wa.me/5511948021428
"""

st.code(mensagem)
