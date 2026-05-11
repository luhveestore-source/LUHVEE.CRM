import streamlit as st
import pandas as pd

st.set_page_config(page_title="LuhVee Store 💖", layout="wide")

st.title("👠 LuhVee Store - Catálogo Automático")

# =========================
# FUNÇÃO BLINDADA CSV
# =========================
def carregar_csv(file):
    encodings = ["utf-8", "latin1", "cp1252"]
    seps = [",", ";", None]

    for enc in encodings:
        for sep in seps:
            try:
                return pd.read_csv(
                    file,
                    encoding=enc,
                    sep=sep,
                    engine="python"
                )
            except:
                continue
    return None

# =========================
# CARREGAR CSV PRINCIPAL
# =========================
st.subheader("👠 Catálogo de Produtos")

try:
    df = carregar_csv("produtos.csv")

    if df is not None:
        st.success("Catálogo carregado 💖")

        st.dataframe(df, use_container_width=True)

        # catálogo visual
        for _, row in df.iterrows():
            st.markdown("---")

            col1, col2 = st.columns([1, 3])

            with col1:
                if "link" in df.columns:
                    st.image(row.get("link", ""), width=120)

            with col2:
                st.markdown(f"### 👠 {row.get('nome','Sem nome')}")
                st.markdown(f"💰 R$ {row.get('preco','')}")
                st.markdown(f"📦 {row.get('categoria','')}")

                whatsapp = "5511948021428"
                msg = f"Quero esse produto: {row.get('nome','')}"

                link = f"https://wa.me/{whatsapp}?text={msg}"

                st.markdown(f"[💬 Comprar no WhatsApp]({link})")

    else:
        st.error("❌ CSV não pôde ser lido. Verifique formato.")

except Exception as e:
    st.error("Erro ao carregar produtos.csv")
    st.write(e)

# =========================
# UPLOAD BLINDADO
# =========================
st.subheader("📂 Upload de CSV")

upload = st.file_uploader("Envie sua planilha", type=["csv"])

if upload is not None:
    df_upload = carregar_csv(upload)

    if df_upload is not None:
        st.dataframe(df_upload)

        if st.button("Salvar catálogo 💖"):
            df_upload.to_csv("produtos.csv", index=False)
            st.success("Atualizado com sucesso!")
    else:
        st.error("❌ Não consegui ler esse CSV")