import streamlit as st
import sqlite3
import pandas as pd
from datetime import datetime

# =====================
# BANCO
# =====================
conn = sqlite3.connect("luhvee.db", check_same_thread=False)
c = conn.cursor()

c.execute("""
CREATE TABLE IF NOT EXISTS leads (
id INTEGER PRIMARY KEY AUTOINCREMENT,
nome TEXT,
whatsapp TEXT,
email TEXT,
interesse TEXT,
data TEXT
)
""")

c.execute("""
CREATE TABLE IF NOT EXISTS produtos (
id INTEGER PRIMARY KEY AUTOINCREMENT,
nome TEXT,
preco REAL,
categoria TEXT,
link TEXT
)
""")

conn.commit()

# =====================
# FUNÇÕES
# =====================
def add_lead(n,w,e,i):
    c.execute("INSERT INTO leads VALUES (NULL,?,?,?,?,?)",
              (n,w,e,i,str(datetime.now())))
    conn.commit()

def get_leads():
    return pd.read_sql("SELECT * FROM leads", conn)

def add_prod(nome,preco,cat,link):
    c.execute("INSERT INTO produtos VALUES (NULL,?,?,?,?)",
              (nome,preco,cat,link))
    conn.commit()

def get_prod():
    return pd.read_sql("SELECT * FROM produtos", conn)

# =====================
# LOGIN SIMPLES
# =====================
st.sidebar.title("🔐 LuhVee Admin")
senha = st.sidebar.text_input("Senha", type="password")

if senha != "luhvee123":
    st.warning("Digite a senha para acessar")
    st.stop()

# =====================
# MENU
# =====================
menu = st.sidebar.selectbox("Menu", ["CRM","Produtos","Dashboard"])

# =====================
# CRM
# =====================
if menu == "CRM":
    st.title("💖 Leads LuhVee")

    n = st.text_input("Nome")
    w = st.text_input("WhatsApp")
    e = st.text_input("Email")
    i = st.selectbox("Interesse",["Sapatos","Botas","Scarpin","Sandálias"])

    if st.button("Salvar"):
        add_lead(n,w,e,i)
        st.success("Salvo 💖")

        st.code(f"""
Olá 💖
Temos novidades em {i} 👠✨

Clique aqui:
https://wa.me/55{w}
""")

    st.dataframe(get_leads())

# =====================
# PRODUTOS
# =====================
elif menu == "Produtos":
    st.title("👠 Catálogo LuhVee")

    nome = st.text_input("Nome do produto")
    preco = st.number_input("Preço")
    cat = st.selectbox("Categoria",["Sapatos","Botas","Scarpin"])
    link = st.text_input("Link imagem")

    if st.button("Adicionar"):
        add_prod(nome,preco,cat,link)
        st.success("Produto adicionado")

    st.dataframe(get_prod())

# =====================
# DASHBOARD
# =====================
elif menu == "Dashboard":
    st.title("📊 Dashboard")

    leads = get_leads()
    prod = get_prod()

    st.metric("Leads", len(leads))
    st.metric("Produtos", len(prod))

    st.bar_chart(leads["interesse"].value_counts())
