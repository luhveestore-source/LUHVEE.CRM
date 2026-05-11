import streamlit as st
import pandas as pd

st.set_page_config(page_title="LuhVee CRM 💖", layout="wide")

st.title("💖 LuhVee Store - CRM + Catálogo Automático 👠")

# =========================
# CARREGAR CSV DE PRODUTOS
# =========================

st.subheader("👠 Catálogo de Produtos")

try:
    df = pd.read_csv("produtos.csv")

    st.success("Produtos carregados com sucesso 💖")

    # mostrar tabela
    st.dataframe(df, use_container_width=True)

    # catálogo visual
    for i, row in df.iterrows():

        st.markdown("---")

        col1, col2 = st.columns([1, 3])

        with col1:
            if "link" in df.columns:
                st.image(row["link"], width=120)

        with col2:
            st.markdown(f"### 👠 {row['nome']}")
            st.markdown(f"💰 Preço: R$ {row['preco']}")
            st.markdown(f"📦 Categoria: {row['categoria']}")

            whatsapp = "5511948021428"

            msg = f"Olá 💖 tenho interesse no produto: {row['nome']}"

            link_whats = f"https://wa.me/{whatsapp}?text={msg}"

            st.markdown(f"[💬 Quero esse no WhatsApp]({link_whats})")

except Exception as e:
    st.error("❌ Não foi possível carregar o produtos.csv")
    st.write(e)

# =========================
# UPLOAD CSV (backup)
# =========================

st.subheader("📂 Atualizar catálogo via CSV")

upload = st.file_uploader("Envie seu produtos.csv", type=["csv"])

if upload is not None:
    df_upload = pd.read_csv(upload)
    st.dataframe(df_upload)

    if st.button("Salvar atualização 💖"):
        df_upload.to_csv("produtos.csv", index=False)
        st.success("Catálogo atualizado com sucesso!")