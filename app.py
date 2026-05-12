import streamlit as st
import pandas as pd

# =====================================
# CONFIG
# =====================================

st.set_page_config(
    page_title="LuhVee CRM PRO 💖",
    layout="wide"
)

# =====================================
# VISUAL
# =====================================

st.title("💖 LuhVee CRM PRO")
st.caption("CRM + Leads + Catálogo + WhatsApp Automático 👠")

# =====================================
# WHATSAPP
# =====================================

WHATSAPP = "5511948021428"

# =====================================
# FUNÇÃO LEITURA
# =====================================

@st.cache_data

def carregar_arquivo(nome):

    encodings = ["utf-8", "latin1", "cp1252"]

    for enc in encodings:
        try:
            if nome.endswith(".csv"):
                return pd.read_csv(
                    nome,
                    encoding=enc,
                    sep=None,
                    engine="python"
                )
            else:
                return pd.read_excel(nome)
        except:
            continue

    return None

# =====================================
# MENU
# =====================================

menu = st.sidebar.radio(
    "💖 Menu",
    [
        "Dashboard",
        "Leads",
        "Catálogo",
        "Campanhas",
        "WhatsApp"
    ]
)

# =====================================
# DASHBOARD
# =====================================

if menu == "Dashboard":

    st.subheader("📊 Painel Geral")

    try:
        leads = carregar_arquivo("brasi.xlsx")

        total_leads = len(leads)

    except:
        total_leads = 0

    try:
        produtos = carregar_arquivo("produtos.csv")

        total_produtos = len(produtos)

    except:
        total_produtos = 0

    col1, col2 = st.columns(2)

    with col1:
        st.metric("👥 Leads", total_leads)

    with col2:
        st.metric("👠 Produtos", total_produtos)

    st.success("💖 Sistema funcionando!")

# =====================================
# LEADS
# =====================================

if menu == "Leads":

    st.subheader("👥 Base de Leads")

    try:

        leads = carregar_arquivo("brasi.xlsx")

        st.success("💖 Leads carregados")

        st.dataframe(leads, use_container_width=True)

        csv = leads.to_csv(index=False).encode("utf-8")

        st.download_button(
            "⬇️ Baixar Leads",
            csv,
            "leads_luhvee.csv",
            "text/csv"
        )

    except Exception as e:
        st.error("❌ Não encontrei brasi.xlsx")
        st.write(e)

# =====================================
# CATÁLOGO
# =====================================

if menu == "Catálogo":

    st.subheader("👠 Catálogo LuhVee")

    try:

        produtos = carregar_arquivo("produtos.csv")

        st.success("💖 Catálogo carregado")

        st.dataframe(produtos, use_container_width=True)

        for _, row in produtos.iterrows():

            st.markdown("---")

            col1, col2 = st.columns([1,3])

            with col1:

                if "link" in produtos.columns:
                    st.image(row.get("link", ""), width=130)

            with col2:

                nome = row.get("nome", "Produto")
                preco = row.get("preco", "")
                categoria = row.get("categoria", "")

                st.markdown(f"### 👠 {nome}")
                st.write(f"💰 R$ {preco}")
                st.write(f"📦 {categoria}")

                mensagem = f'''💖 Olá!\n\nTenho interesse neste produto:\n\n👠 {nome}\n💰 R$ {preco}\n\nPode me enviar mais detalhes?'''

                link = f"https://wa.me/{WHATSAPP}?text={mensagem}"

                st.markdown(f"[💬 Comprar Agora]({link})")

    except Exception as e:
        st.error("❌ produtos.csv não encontrado")
        st.write(e)

# =====================================
# CAMPANHAS
# =====================================

if menu == "Campanhas":

    st.subheader("📢 Campanhas Prontas")

    mensagem1 = """💖 Olá! Tudo bem?\n\nA LuhVee Stores acabou de receber novidades incríveis 👠✨\n\nTemos modelos lindos com preços especiais hoje 💕\n\nQuer receber o catálogo?"""

    mensagem2 = """🔥 PROMOÇÃO LUHVEE STORES 🔥\n\nSapatos incríveis com preços especiais 👠💕\n\nMe chama agora e garanta o seu antes que acabe ✨"""

    mensagem3 = """💖 Novidades chegando na LuhVee Stores 👠✨\n\nSelecionamos modelos perfeitos pra você!\n\nQuer ver as novidades disponíveis hoje?"""

    st.text_area("💬 Campanha 1", mensagem1, height=180)

    st.text_area("💬 Campanha 2", mensagem2, height=180)

    st.text_area("💬 Campanha 3", mensagem3, height=180)

# =====================================
# WHATSAPP
# =====================================

if menu == "WhatsApp":

    st.subheader("💬 Central WhatsApp")

    mensagem = st.text_area(
        "Mensagem",
        "💖 Olá! Conheça as novidades da LuhVee Stores 👠✨"
    )

    link = f"https://wa.me/{WHATSAPP}?text={mensagem}"

    st.markdown(f"[🚀 Abrir WhatsApp]({link})")

    st.info("💡 Use campanhas menores para evitar bloqueios no WhatsApp.")