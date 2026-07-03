import streamlit as st
import pandas as pd
import os
import re
import urllib.parse
from datetime import datetime

# ==========================================================
# CONFIGURAÇÃO GERAL
# ==========================================================

st.set_page_config(
    page_title="LuhVee CRM PRO 💖",
    layout="wide"
)

WHATSAPP_LOJA = "5511948021428"

ARQUIVO_BASE = "crm_luhvee_base.csv"
ARQUIVO_ANTIGO = "brasi.xlsx"
ARQUIVO_SP = "sp.xlsx"
ARQUIVO_PRODUTOS = "produtos.csv"

STATUS_OPCOES = [
    "Novo",
    "Contato feito",
    "Respondeu",
    "Interessado",
    "Enviar catálogo",
    "Aguardando retorno",
    "Cliente fechado",
    "Não respondeu",
    "Não tem interesse",
    "Bloquear/Não contatar"
]

TIPO_CLIENTE_OPCOES = [
    "Pessoa Física",
    "MEI/CNPJ",
    "Não informado"
]

# ==========================================================
# VISUAL
# ==========================================================

st.title("💖 LuhVee CRM PRO")
st.caption("CRM + Leads + Prospecção + Catálogo + Mensagens WhatsApp/E-mail ✨")

# ==========================================================
# FUNÇÕES AUXILIARES
# ==========================================================

def limpar_texto(valor):
    if pd.isna(valor):
        return ""
    valor = str(valor).strip()
    if valor.lower() in ["nan", "none", "null"]:
        return ""
    return valor


def somente_numeros(valor):
    return re.sub(r"\D", "", limpar_texto(valor))


def formatar_telefone(valor):
    tel = somente_numeros(valor)

    if not tel:
        return ""

    # Remove zeros no começo
    tel = tel.lstrip("0")

    # Se vier sem DDI e tiver 10 ou 11 dígitos, adiciona 55
    if len(tel) in [10, 11]:
        tel = "55" + tel

    return tel


def primeiro_nome(nome):
    nome = limpar_texto(nome)
    if not nome:
        return "tudo bem"
    partes = nome.split()
    if len(partes) <= 1:
        return nome.title()
    # Em bases compradas pode vir CPF/CNPJ ou código no começo.
    for parte in partes:
        if not any(char.isdigit() for char in parte) and len(parte) > 2:
            return parte.title()
    return partes[0].title()


def detectar_tipo_cliente(nome, documento=""):
    texto = f"{limpar_texto(nome)} {limpar_texto(documento)}".upper()
    numeros = somente_numeros(texto)

    # CNPJ tem 14 dígitos; CPF tem 11.
    if len(numeros) >= 14:
        return "MEI/CNPJ"

    termos_empresa = [
        "LTDA", "ME", "MEI", "EIRELI", "COMERCIO", "COMÉRCIO",
        "SERVICOS", "SERVIÇOS", "EMPRESA", "CNPJ"
    ]

    if any(t in texto for t in termos_empresa):
        return "MEI/CNPJ"

    return "Pessoa Física"


def carregar_arquivo(nome):
    if not os.path.exists(nome):
        return None

    encodings = ["utf-8", "latin1", "cp1252"]

    for enc in encodings:
        try:
            if nome.lower().endswith(".csv"):
                return pd.read_csv(nome, encoding=enc, sep=None, engine="python")
            return pd.read_excel(nome)
        except Exception:
            continue

    return None


def padronizar_sp(df):
    """
    Padroniza a planilha sp.xlsx.
    A planilha veio sem cabeçalho correto, então usamos posição das colunas.
    """
    if df is None or df.empty:
        return pd.DataFrame()

    df = df.copy()

    # Se a primeira linha virou cabeçalho por engano, recuperamos o cabeçalho como linha.
    primeira_linha = pd.DataFrame([list(df.columns)], columns=df.columns)
    df = pd.concat([primeira_linha, df], ignore_index=True)

    colunas = list(df.columns)

    def pegar_coluna(pos):
        if len(colunas) > pos:
            return df[colunas[pos]].apply(limpar_texto)
        return ""

    base = pd.DataFrame()
    base["Nome"] = pegar_coluna(0)
    base["Documento"] = ""
    base["Tipo Cliente"] = "Pessoa Física"
    base["Tipo Logradouro"] = pegar_coluna(2)
    base["Endereço"] = pegar_coluna(3)
    base["Número"] = pegar_coluna(4)
    base["Complemento"] = pegar_coluna(5)
    base["Bairro"] = pegar_coluna(6)
    base["Cidade"] = pegar_coluna(7)
    base["UF"] = pegar_coluna(8)
    base["CEP"] = pegar_coluna(10)
    base["Telefone"] = pegar_coluna(11).apply(formatar_telefone)
    base["Email"] = pegar_coluna(13)
    base["Site"] = ""
    base["Origem"] = "Lista SP Pessoa Física"
    base["Status"] = "Novo"
    base["Interesse"] = ""
    base["Observações"] = ""
    base["Último Contato"] = ""
    base["Data Cadastro"] = datetime.now().strftime("%d/%m/%Y")

    # Remove linhas completamente vazias
    base = base[base["Nome"].str.strip() != ""]

    return base


def padronizar_generico(df, origem="Base antiga"):
    """
    Tenta padronizar qualquer planilha com nomes de colunas variados.
    """
    if df is None or df.empty:
        return pd.DataFrame()

    df = df.copy()
    mapa = {str(c).strip().lower(): c for c in df.columns}

    def buscar_coluna(possiveis):
        for p in possiveis:
            for nome_coluna, original in mapa.items():
                if p in nome_coluna:
                    return df[original].apply(limpar_texto)
        return pd.Series([""] * len(df))

    nome = buscar_coluna(["nome", "cliente", "empresa", "razao", "razão"])
    documento = buscar_coluna(["cnpj", "cpf", "documento"])
    telefone = buscar_coluna(["telefone", "celular", "whatsapp", "fone"])
    email = buscar_coluna(["email", "e-mail", "mail"])
    cidade = buscar_coluna(["cidade", "municipio", "município"])
    uf = buscar_coluna(["uf", "estado"])
    bairro = buscar_coluna(["bairro"])
    endereco = buscar_coluna(["endereco", "endereço", "logradouro", "rua"])
    cep = buscar_coluna(["cep"])
    site = buscar_coluna(["site", "website"])

    base = pd.DataFrame()
    base["Nome"] = nome
    base["Documento"] = documento
    base["Tipo Cliente"] = [
        detectar_tipo_cliente(n, d) for n, d in zip(nome, documento)
    ]
    base["Tipo Logradouro"] = ""
    base["Endereço"] = endereco
    base["Número"] = ""
    base["Complemento"] = ""
    base["Bairro"] = bairro
    base["Cidade"] = cidade
    base["UF"] = uf
    base["CEP"] = cep
    base["Telefone"] = telefone.apply(formatar_telefone)
    base["Email"] = email
    base["Site"] = site
    base["Origem"] = origem
    base["Status"] = "Novo"
    base["Interesse"] = ""
    base["Observações"] = ""
    base["Último Contato"] = ""
    base["Data Cadastro"] = datetime.now().strftime("%d/%m/%Y")

    base = base[base["Nome"].str.strip() != ""]

    return base


def remover_duplicados(df):
    if df is None or df.empty:
        return pd.DataFrame()

    df = df.copy()

    for coluna in ["Telefone", "Email", "Nome"]:
        if coluna not in df.columns:
            df[coluna] = ""

    df["Telefone"] = df["Telefone"].apply(formatar_telefone)
    df["Email"] = df["Email"].apply(lambda x: limpar_texto(x).lower())

    # Chave de duplicidade: telefone primeiro, depois e-mail, depois nome+cidade.
    df["chave_temp"] = df.apply(
        lambda r: r["Telefone"]
        if r["Telefone"]
        else (
            r["Email"]
            if r["Email"]
            else f'{limpar_texto(r.get("Nome", "")).lower()}_{limpar_texto(r.get("Cidade", "")).lower()}'
        ),
        axis=1
    )

    df = df.drop_duplicates(subset=["chave_temp"], keep="first")
    df = df.drop(columns=["chave_temp"], errors="ignore")

    return df


def montar_base_inicial():
    bases = []

    antiga = carregar_arquivo(ARQUIVO_ANTIGO)
    if antiga is not None:
        bases.append(padronizar_generico(antiga, "Base antiga MEI/CNPJ"))

    sp = carregar_arquivo(ARQUIVO_SP)
    if sp is not None:
        bases.append(padronizar_sp(sp))

    if not bases:
        return pd.DataFrame(columns=[
            "Nome", "Documento", "Tipo Cliente", "Tipo Logradouro", "Endereço",
            "Número", "Complemento", "Bairro", "Cidade", "UF", "CEP",
            "Telefone", "Email", "Site", "Origem", "Status", "Interesse",
            "Observações", "Último Contato", "Data Cadastro"
        ])

    base = pd.concat(bases, ignore_index=True)
    base = remover_duplicados(base)

    return base


def carregar_base_crm():
    if os.path.exists(ARQUIVO_BASE):
        try:
            return pd.read_csv(ARQUIVO_BASE, dtype=str).fillna("")
        except Exception:
            pass

    base = montar_base_inicial()
    salvar_base_crm(base)
    return base


def salvar_base_crm(df):
    df.to_csv(ARQUIVO_BASE, index=False, encoding="utf-8-sig")


def gerar_msg_whatsapp(row, tom="acolhedora"):
    nome = primeiro_nome(row.get("Nome", ""))
    tipo = row.get("Tipo Cliente", "Pessoa Física")

    if tom == "curta":
        return f"""Oi, {nome}! Tudo bem? 💖

Sou da LuhVee Stores ❤️
Trabalhamos com perfumes árabes originais, cosméticos, body splash, hidratantes, kits para presente e achadinhos selecionados.

Posso te enviar nosso catálogo com as novidades?"""

    if tipo == "MEI/CNPJ":
        return f"""Olá, {nome}! Tudo bem? 💖

Sou da LuhVee Stores ❤️
Trabalhamos com curadoria de produtos para beleza, autocuidado, presentes, perfumes árabes originais, cosméticos, body splash e kits especiais.

Gostaria de apresentar algumas opções que podem ser interessantes para você, sua empresa ou para presentear clientes e colaboradores.

Posso te enviar nosso catálogo com as novidades?"""

    return f"""Oi, {nome}! Tudo bem? 💖

Sou da LuhVee Stores ❤️
Uma loja feita com carinho para quem ama se cuidar, presentear e encontrar achadinhos especiais.

Temos perfumes árabes originais, cosméticos, body splash, hidratantes, kits para presente e diversos produtos selecionados com muito cuidado.

Posso te enviar nosso catálogo com as novidades disponíveis?"""


def gerar_msg_email(row):
    nome = primeiro_nome(row.get("Nome", ""))

    return f"""Olá, {nome}! Tudo bem?

Prazer, somos a LuhVee Stores ❤️

A LuhVee Stores é uma loja criada com carinho para oferecer produtos selecionados para autocuidado, presentes e achadinhos especiais.

Temos perfumes árabes originais, cosméticos, body splash, hidratantes, kits para presente e diversos produtos para quem ama se cuidar ou surpreender alguém especial.

Será um prazer te apresentar nossas novidades e opções disponíveis.

Caso queira receber nosso catálogo, é só responder este e-mail ou chamar no WhatsApp:
https://wa.me/{WHATSAPP_LOJA}

Com carinho,
LuhVee Stores ❤️"""


def gerar_link_whatsapp(telefone, mensagem):
    telefone = formatar_telefone(telefone)
    texto = urllib.parse.quote(mensagem)

    if telefone:
        return f"https://wa.me/{telefone}?text={texto}"

    return f"https://wa.me/{WHATSAPP_LOJA}?text={texto}"


def aplicar_filtros(df):
    filtrado = df.copy()

    col1, col2, col3, col4 = st.columns(4)

    with col1:
        tipo = st.selectbox("Tipo de cliente", ["Todos"] + TIPO_CLIENTE_OPCOES)

    with col2:
        status = st.selectbox("Status", ["Todos"] + STATUS_OPCOES)

    with col3:
        origem = st.selectbox("Origem", ["Todas"] + sorted(filtrado["Origem"].dropna().unique().tolist()))

    with col4:
        busca = st.text_input("Buscar por nome, telefone, e-mail ou bairro")

    if tipo != "Todos":
        filtrado = filtrado[filtrado["Tipo Cliente"] == tipo]

    if status != "Todos":
        filtrado = filtrado[filtrado["Status"] == status]

    if origem != "Todas":
        filtrado = filtrado[filtrado["Origem"] == origem]

    if busca:
        busca = busca.lower()
        texto = filtrado.astype(str).agg(" ".join, axis=1).str.lower()
        filtrado = filtrado[texto.str.contains(busca, na=False)]

    return filtrado

# ==========================================================
# MENU
# ==========================================================

menu = st.sidebar.radio(
    "💖 Menu",
    [
        "Dashboard",
        "Leads",
        "Adicionar/Importar",
        "Atualizar Contato",
        "Mensagens",
        "Catálogo",
        "Campanhas",
        "WhatsApp"
    ]
)

# ==========================================================
# DASHBOARD
# ==========================================================

if menu == "Dashboard":
    st.subheader("📊 Painel Geral")

    leads = carregar_base_crm()

    total_leads = len(leads)
    pf = len(leads[leads["Tipo Cliente"] == "Pessoa Física"])
    pj = len(leads[leads["Tipo Cliente"] == "MEI/CNPJ"])

    try:
        produtos = carregar_arquivo(ARQUIVO_PRODUTOS)
        total_produtos = len(produtos) if produtos is not None else 0
    except Exception:
        total_produtos = 0

    col1, col2, col3, col4 = st.columns(4)

    col1.metric("👥 Total Leads", total_leads)
    col2.metric("🙋 Pessoa Física", pf)
    col3.metric("🏢 MEI/CNPJ", pj)
    col4.metric("👠 Produtos", total_produtos)

    st.markdown("---")

    col5, col6 = st.columns(2)

    with col5:
        st.subheader("Status dos Leads")
        if not leads.empty:
            st.bar_chart(leads["Status"].value_counts())

    with col6:
        st.subheader("Origem dos Leads")
        if not leads.empty:
            st.bar_chart(leads["Origem"].value_counts())

    st.info("💡 Dica LGPD: envie mensagens personalizadas, com identificação da loja e opção da pessoa pedir para não receber novos contatos.")

# ==========================================================
# LEADS
# ==========================================================

if menu == "Leads":
    st.subheader("👥 Base de Leads")

    leads = carregar_base_crm()

    if leads.empty:
        st.warning("Nenhum lead encontrado ainda.")
    else:
        filtrado = aplicar_filtros(leads)

        st.success(f"💖 {len(filtrado)} leads encontrados")

        st.dataframe(filtrado, use_container_width=True, height=520)

        csv = filtrado.to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            "⬇️ Baixar Leads Filtrados",
            csv,
            "leads_luhvee_filtrados.csv",
            "text/csv"
        )

        st.download_button(
            "⬇️ Baixar Base Completa",
            leads.to_csv(index=False).encode("utf-8-sig"),
            "crm_luhvee_base_completa.csv",
            "text/csv"
        )

# ==========================================================
# ADICIONAR / IMPORTAR
# ==========================================================

if menu == "Adicionar/Importar":
    st.subheader("➕ Adicionar ou importar leads")

    leads = carregar_base_crm()

    st.markdown("### Importar nova lista Excel/CSV")
    arquivo = st.file_uploader("Enviar arquivo de leads", type=["xlsx", "xls", "csv"])

    tipo_importacao = st.selectbox(
        "Tipo da lista importada",
        ["Pessoa Física", "MEI/CNPJ", "Detectar automaticamente"]
    )

    origem_importacao = st.text_input("Origem da lista", "Nova lista importada")

    if arquivo is not None:
        try:
            if arquivo.name.lower().endswith(".csv"):
                novo_df = pd.read_csv(arquivo, sep=None, engine="python", dtype=str).fillna("")
            else:
                novo_df = pd.read_excel(arquivo, dtype=str).fillna("")

            base_nova = padronizar_generico(novo_df, origem_importacao)

            if tipo_importacao != "Detectar automaticamente":
                base_nova["Tipo Cliente"] = tipo_importacao

            antes = len(leads)
            atualizado = pd.concat([leads, base_nova], ignore_index=True)
            atualizado = remover_duplicados(atualizado)
            depois = len(atualizado)

            st.write("Prévia da importação:")
            st.dataframe(base_nova.head(20), use_container_width=True)

            if st.button("✅ Confirmar importação"):
                salvar_base_crm(atualizado)
                st.success(f"Importação concluída! Base tinha {antes} leads e agora ficou com {depois} leads sem duplicados.")
                st.rerun()

        except Exception as e:
            st.error("Não consegui importar esse arquivo.")
            st.write(e)

    st.markdown("---")
    st.markdown("### Cadastro manual")

    with st.form("cadastro_manual"):
        col1, col2, col3 = st.columns(3)
        nome = col1.text_input("Nome")
        telefone = col2.text_input("WhatsApp/Telefone")
        email = col3.text_input("E-mail")

        col4, col5, col6 = st.columns(3)
        tipo_cliente = col4.selectbox("Tipo Cliente", TIPO_CLIENTE_OPCOES)
        cidade = col5.text_input("Cidade")
        bairro = col6.text_input("Bairro")

        endereco = st.text_input("Endereço")
        interesse = st.text_input("Interesse")
        obs = st.text_area("Observações")

        salvar = st.form_submit_button("💾 Salvar lead")

    if salvar:
        novo = pd.DataFrame([{
            "Nome": nome,
            "Documento": "",
            "Tipo Cliente": tipo_cliente,
            "Tipo Logradouro": "",
            "Endereço": endereco,
            "Número": "",
            "Complemento": "",
            "Bairro": bairro,
            "Cidade": cidade,
            "UF": "",
            "CEP": "",
            "Telefone": formatar_telefone(telefone),
            "Email": email,
            "Site": "",
            "Origem": "Cadastro manual",
            "Status": "Novo",
            "Interesse": interesse,
            "Observações": obs,
            "Último Contato": "",
            "Data Cadastro": datetime.now().strftime("%d/%m/%Y")
        }])

        leads = pd.concat([leads, novo], ignore_index=True)
        leads = remover_duplicados(leads)
        salvar_base_crm(leads)
        st.success("Lead salvo com sucesso!")
        st.rerun()

# ==========================================================
# ATUALIZAR CONTATO
# ==========================================================

if menu == "Atualizar Contato":
    st.subheader("📝 Atualizar status e observações")

    leads = carregar_base_crm()

    if leads.empty:
        st.warning("Nenhum lead cadastrado.")
    else:
        busca = st.text_input("Buscar lead por nome, telefone ou e-mail")

        if busca:
            texto = leads.astype(str).agg(" ".join, axis=1).str.lower()
            resultados = leads[texto.str.contains(busca.lower(), na=False)]
        else:
            resultados = leads.head(50)

        st.dataframe(resultados, use_container_width=True, height=260)

        if not resultados.empty:
            opcoes = resultados.index.tolist()
            idx = st.selectbox(
                "Escolha o lead para atualizar",
                opcoes,
                format_func=lambda i: f'{leads.loc[i, "Nome"]} | {leads.loc[i, "Telefone"]} | {leads.loc[i, "Status"]}'
            )

            with st.form("editar_lead"):
                col1, col2, col3 = st.columns(3)
                status = col1.selectbox("Status", STATUS_OPCOES, index=STATUS_OPCOES.index(leads.loc[idx, "Status"]) if leads.loc[idx, "Status"] in STATUS_OPCOES else 0)
                interesse = col2.text_input("Interesse", leads.loc[idx, "Interesse"])
                ultimo = col3.text_input("Último Contato", datetime.now().strftime("%d/%m/%Y"))

                obs = st.text_area("Observações", leads.loc[idx, "Observações"])

                salvar_edicao = st.form_submit_button("💾 Atualizar")

            if salvar_edicao:
                leads.loc[idx, "Status"] = status
                leads.loc[idx, "Interesse"] = interesse
                leads.loc[idx, "Último Contato"] = ultimo
                leads.loc[idx, "Observações"] = obs
                salvar_base_crm(leads)
                st.success("Contato atualizado!")
                st.rerun()

# ==========================================================
# MENSAGENS
# ==========================================================

if menu == "Mensagens":
    st.subheader("💌 Gerador de mensagens por lead")

    leads = carregar_base_crm()

    if leads.empty:
        st.warning("Nenhum lead cadastrado.")
    else:
        filtrado = aplicar_filtros(leads)

        if filtrado.empty:
            st.warning("Nenhum lead encontrado nesse filtro.")
        else:
            idx = st.selectbox(
                "Escolha o lead",
                filtrado.index.tolist(),
                format_func=lambda i: f'{leads.loc[i, "Nome"]} | {leads.loc[i, "Telefone"]} | {leads.loc[i, "Email"]}'
            )

            row = leads.loc[idx]

            tom = st.radio("Modelo de WhatsApp", ["acolhedora", "curta"], horizontal=True)

            msg_wpp = gerar_msg_whatsapp(row, tom)
            msg_email = gerar_msg_email(row)

            st.markdown("### WhatsApp")
            msg_wpp_editada = st.text_area("Mensagem para WhatsApp", msg_wpp, height=230)

            link = gerar_link_whatsapp(row.get("Telefone", ""), msg_wpp_editada)
            st.markdown(f"[🚀 Abrir WhatsApp do lead]({link})")

            st.markdown("### E-mail")
            assunto = "Conheça a LuhVee Stores ❤️"
            st.text_input("Assunto", assunto)
            st.text_area("Mensagem para e-mail", msg_email, height=300)

            email = limpar_texto(row.get("Email", ""))
            if email:
                mailto = f"mailto:{email}?subject={urllib.parse.quote(assunto)}&body={urllib.parse.quote(msg_email)}"
                st.markdown(f"[📧 Abrir e-mail para o lead]({mailto})")
            else:
                st.info("Esse lead não tem e-mail cadastrado.")

            st.warning("⚠️ Evite disparo em massa. Personalize a abordagem e respeite pedidos de remoção da lista.")

# ==========================================================
# CATÁLOGO
# ==========================================================

if menu == "Catálogo":
    st.subheader("👠 Catálogo LuhVee")

    try:
        produtos = carregar_arquivo(ARQUIVO_PRODUTOS)

        if produtos is None or produtos.empty:
            st.warning("produtos.csv não encontrado ou vazio.")
        else:
            st.success("💖 Catálogo carregado")
            st.dataframe(produtos, use_container_width=True)

            for _, row in produtos.iterrows():
                st.markdown("---")
                col1, col2 = st.columns([1, 3])

                with col1:
                    link_img = row.get("link", "") if "link" in produtos.columns else ""
                    if link_img:
                        try:
                            st.image(link_img, width=130)
                        except Exception:
                            st.write("Sem imagem")

                with col2:
                    nome = row.get("nome", "Produto")
                    preco = row.get("preco", "")
                    categoria = row.get("categoria", "")

                    st.markdown(f"### 💖 {nome}")
                    st.write(f"💰 R$ {preco}")
                    st.write(f"📦 {categoria}")

                    mensagem = f"""💖 Olá!

Tenho interesse neste produto da LuhVee Stores:

✨ {nome}
💰 R$ {preco}

Pode me enviar mais detalhes?"""

                    link = f"https://wa.me/{WHATSAPP_LOJA}?text={urllib.parse.quote(mensagem)}"
                    st.markdown(f"[💬 Comprar Agora]({link})")

    except Exception as e:
        st.error("❌ Erro ao carregar catálogo")
        st.write(e)

# ==========================================================
# CAMPANHAS
# ==========================================================

if menu == "Campanhas":
    st.subheader("📢 Campanhas prontas")

    mensagem1 = """💖 Olá! Tudo bem?

A LuhVee Stores preparou novidades especiais para quem ama se cuidar e presentear ✨

Temos perfumes árabes originais, cosméticos, body splash, hidratantes, kits e achadinhos selecionados com carinho.

Quer receber nosso catálogo?"""

    mensagem2 = """✨ Novidades LuhVee Stores ❤️

Produtos para autocuidado, presentes e aquele mimo especial do dia a dia.

Perfumes árabes originais, cosméticos, body splash, hidratantes e muito mais.

Me chama que eu te envio as opções disponíveis 💖"""

    mensagem3 = """💖 Presenteie com carinho!

A LuhVee Stores monta kits lindos e personalizados para várias ocasiões 🎁

Temos perfumes árabes originais, hidratantes, body splash, cosméticos e diversos produtos especiais.

Quer ver as opções de hoje?"""

    mensagem_lgpd = """Olá! Sou da LuhVee Stores ❤️

Estou entrando em contato para apresentar nossos produtos de autocuidado, presentes e achadinhos especiais.

Caso não queira receber novas mensagens, é só me avisar que removo seu contato da nossa lista. 💖"""

    st.text_area("💬 Campanha 1 - Apresentação", mensagem1, height=180)
    st.text_area("💬 Campanha 2 - Novidades", mensagem2, height=180)
    st.text_area("💬 Campanha 3 - Kits e presentes", mensagem3, height=180)
    st.text_area("💬 Rodapé opcional LGPD", mensagem_lgpd, height=150)

# ==========================================================
# WHATSAPP
# ==========================================================

if menu == "WhatsApp":
    st.subheader("💬 Central WhatsApp")

    mensagem = st.text_area(
        "Mensagem",
        "💖 Olá! Conheça as novidades da LuhVee Stores ❤️\n\nTemos perfumes árabes originais, cosméticos, body splash, kits e achadinhos especiais."
    )

    link = f"https://wa.me/{WHATSAPP_LOJA}?text={urllib.parse.quote(mensagem)}"
    st.markdown(f"[🚀 Abrir WhatsApp da loja]({link})")

    st.info("💡 Use mensagens menores, personalizadas e evite disparos em massa para não bloquear o WhatsApp.")
