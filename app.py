import streamlit as st
import pandas as pd
import os
import re
import urllib.parse
from io import StringIO
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
ARQUIVO_ESTOQUE = "estoque_base.xlsx"
ARQUIVO_PRODUTOS_CSV = "produtos.csv"

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
st.caption("CRM + Leads + Estoque + Catálogo + WhatsApp/E-mail Automático ✨")


# ==========================================================
# FUNÇÕES GERAIS
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

    tel = tel.lstrip("0")

    if len(tel) in [10, 11]:
        tel = "55" + tel

    return tel


def converter_numero(valor):
    valor = limpar_texto(valor)

    if not valor:
        return 0.0

    valor = valor.replace("R$", "").replace(" ", "")

    # Formato brasileiro: 1.234,56
    if "," in valor:
        valor = valor.replace(".", "").replace(",", ".")

    try:
        return float(valor)
    except Exception:
        return 0.0


def moeda(valor):
    valor = converter_numero(valor)
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def primeiro_nome(nome):
    nome = limpar_texto(nome)
    if not nome:
        return "tudo bem"

    partes = nome.split()
    for parte in partes:
        if not any(char.isdigit() for char in parte) and len(parte) > 2:
            return parte.title()

    return partes[0].title()


def detectar_tipo_cliente(nome, documento=""):
    texto = f"{limpar_texto(nome)} {limpar_texto(documento)}".upper()
    numeros = somente_numeros(texto)

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


# ==========================================================
# LEADS
# ==========================================================

def padronizar_sp(df):
    """
    Padroniza a planilha sp.xlsx.
    A planilha veio sem cabeçalho correto, então usamos a posição das colunas.
    """
    if df is None or df.empty:
        return pd.DataFrame()

    df = df.copy()

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

    base = base[base["Nome"].str.strip() != ""]

    return base


def padronizar_generico(df, origem="Base antiga"):
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

    for coluna in ["Telefone", "Email", "Nome", "Cidade"]:
        if coluna not in df.columns:
            df[coluna] = ""

    df["Telefone"] = df["Telefone"].apply(formatar_telefone)
    df["Email"] = df["Email"].apply(lambda x: limpar_texto(x).lower())

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


# ==========================================================
# ESTOQUE / CATÁLOGO
# ==========================================================

def carregar_produtos():
    """
    Carrega estoque_base.xlsx ou produtos.csv.
    A planilha estoque_base.xlsx pode vir em uma única coluna separada por vírgulas.
    O sistema mostra para o cliente somente: Produto, Categoria, Preço Venda e Estoque.
    """
    df = None

    if os.path.exists(ARQUIVO_ESTOQUE):
        df = carregar_arquivo(ARQUIVO_ESTOQUE)

        if df is not None and not df.empty and len(df.columns) == 1:
            coluna_unica = df.columns[0]
            linhas = [str(coluna_unica)]
            linhas += df[coluna_unica].dropna().astype(str).tolist()
            texto = "\n".join(linhas)
            df = pd.read_csv(StringIO(texto), sep=",", engine="python")

    elif os.path.exists(ARQUIVO_PRODUTOS_CSV):
        df = carregar_arquivo(ARQUIVO_PRODUTOS_CSV)

    if df is None or df.empty:
        return pd.DataFrame(columns=[
            "Código", "Produto", "Categoria", "Fornecedor",
            "Custo", "Preço Venda", "Estoque", "Disponível"
        ])

    renomear = {}

    for c in df.columns:
        c_original = c
        c_limpo = str(c).strip().upper()
        c_limpo = (
            c_limpo
            .replace("CÃ“DIGO", "CÓDIGO")
            .replace("PREÃ‡O", "PREÇO")
            .replace("CÃ‡", "Ç")
        )

        if "CODIGO" in c_limpo or "CÓDIGO" in c_limpo or c_limpo == "COD":
            renomear[c_original] = "Código"
        elif "PRODUTO" in c_limpo or "NOME" in c_limpo:
            renomear[c_original] = "Produto"
        elif "CATEGORIA" in c_limpo:
            renomear[c_original] = "Categoria"
        elif "FORNECEDOR" in c_limpo:
            renomear[c_original] = "Fornecedor"
        elif "CUSTO" in c_limpo:
            renomear[c_original] = "Custo"
        elif "PRECO" in c_limpo or "PREÇO" in c_limpo or "VENDA" in c_limpo or "VALOR" in c_limpo:
            renomear[c_original] = "Preço Venda"
        elif "ESTOQUE" in c_limpo or "QTD" in c_limpo or "QUANTIDADE" in c_limpo:
            renomear[c_original] = "Estoque"
        elif "CATEGORIA" in c_limpo:
            renomear[c_original] = "Categoria"

    df = df.rename(columns=renomear)

    for coluna in ["Código", "Produto", "Categoria", "Fornecedor", "Custo", "Preço Venda", "Estoque"]:
        if coluna not in df.columns:
            df[coluna] = ""

    df["Produto"] = df["Produto"].apply(limpar_texto)
    df["Categoria"] = df["Categoria"].apply(limpar_texto)
    df["Preço Venda"] = df["Preço Venda"].apply(converter_numero)
    df["Estoque"] = df["Estoque"].apply(converter_numero)
    df["Disponível"] = df["Estoque"].apply(lambda x: "Sim" if x > 0 else "Não")

    df = df[df["Produto"] != ""].copy()

    return df


def catalogo_para_cliente(produtos, categoria="Todas", limite=12):
    if produtos is None or produtos.empty:
        return "No momento estou organizando o catálogo da LuhVee Stores ❤️"

    df = produtos.copy()

    if categoria != "Todas":
        df = df[df["Categoria"] == categoria]

    df = df[df["Estoque"] > 0]

    if df.empty:
        return "No momento não encontrei produtos disponíveis nessa categoria."

    linhas = []
    for _, p in df.head(limite).iterrows():
        nome = limpar_texto(p.get("Produto", "Produto"))
        preco = moeda(p.get("Preço Venda", 0))
        categoria_produto = limpar_texto(p.get("Categoria", ""))

        if categoria_produto:
            linhas.append(f"✨ {nome} | {categoria_produto} | {preco}")
        else:
            linhas.append(f"✨ {nome} | {preco}")

    return "\n".join(linhas)


# ==========================================================
# MENSAGENS
# ==========================================================

def gerar_msg_whatsapp(row, produtos=None, categoria="Todas", tom="acolhedora", incluir_produtos=True):
    nome = primeiro_nome(row.get("Nome", ""))
    tipo = row.get("Tipo Cliente", "Pessoa Física")

    lista_produtos = ""
    if incluir_produtos and produtos is not None and not produtos.empty:
        lista_produtos = catalogo_para_cliente(produtos, categoria=categoria, limite=8)

    if tipo == "MEI/CNPJ":
        mensagem = f"""Olá, {nome}! Tudo bem? 💖

Sou da LuhVee Stores ❤️
Trabalhamos com produtos à pronta entrega para autocuidado, presentes, perfumes árabes originais, cosméticos, body splash, hidratantes e kits especiais.

Gostaria de apresentar algumas opções que podem ser interessantes para você, sua empresa ou para presentear clientes e colaboradores."""
    else:
        mensagem = f"""Oi, {nome}! Tudo bem? 💖

Sou da LuhVee Stores ❤️
Uma loja feita com carinho para quem ama se cuidar, presentear e encontrar achadinhos especiais.

Temos produtos à pronta entrega, perfumes árabes originais, cosméticos, body splash, hidratantes, kits e diversos produtos selecionados com cuidado."""

    if incluir_produtos and lista_produtos:
        mensagem += f"""

Separei algumas opções disponíveis hoje:

{lista_produtos}

Quer que eu te envie mais detalhes ou fotos dos produtos?"""
    else:
        mensagem += """

Posso te enviar nosso catálogo com as novidades disponíveis?"""

    if tom == "curta":
        mensagem = f"""Oi, {nome}! Tudo bem? 💖

Sou da LuhVee Stores ❤️
Temos produtos à pronta entrega: perfumes árabes originais, cosméticos, body splash, hidratantes, kits e achadinhos especiais.

Posso te enviar o catálogo?"""

    return mensagem


def gerar_msg_email(row, produtos=None, categoria="Todas", incluir_produtos=True):
    nome = primeiro_nome(row.get("Nome", ""))

    lista_produtos = ""
    if incluir_produtos and produtos is not None and not produtos.empty:
        lista_produtos = catalogo_para_cliente(produtos, categoria=categoria, limite=10)

    mensagem = f"""Olá, {nome}! Tudo bem?

Prazer, somos a LuhVee Stores ❤️

A LuhVee Stores é uma loja criada com carinho para oferecer produtos selecionados para autocuidado, presentes e achadinhos especiais.

Temos produtos à pronta entrega, perfumes árabes originais, cosméticos, body splash, hidratantes, kits para presente e diversos produtos para quem ama se cuidar ou surpreender alguém especial."""

    if incluir_produtos and lista_produtos:
        mensagem += f"""

Algumas opções disponíveis hoje:

{lista_produtos}"""

    mensagem += f"""

Será um prazer te apresentar nossas novidades e opções disponíveis.

Caso queira receber nosso catálogo completo, é só responder este e-mail ou chamar no WhatsApp:
https://wa.me/{WHATSAPP_LOJA}

Com carinho,
LuhVee Stores ❤️"""

    return mensagem


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
        origens = sorted(filtrado["Origem"].dropna().unique().tolist()) if "Origem" in filtrado.columns else []
        origem = st.selectbox("Origem", ["Todas"] + origens)

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
    produtos = carregar_produtos()

    total_leads = len(leads)
    pf = len(leads[leads["Tipo Cliente"] == "Pessoa Física"]) if not leads.empty else 0
    pj = len(leads[leads["Tipo Cliente"] == "MEI/CNPJ"]) if not leads.empty else 0
    total_produtos = len(produtos)
    produtos_disponiveis = len(produtos[produtos["Estoque"] > 0]) if not produtos.empty else 0

    col1, col2, col3, col4, col5 = st.columns(5)

    col1.metric("👥 Total Leads", total_leads)
    col2.metric("🙋 Pessoa Física", pf)
    col3.metric("🏢 MEI/CNPJ", pj)
    col4.metric("🛍️ Produtos", total_produtos)
    col5.metric("✅ À pronta entrega", produtos_disponiveis)

    st.markdown("---")

    col6, col7 = st.columns(2)

    with col6:
        st.subheader("Status dos Leads")
        if not leads.empty:
            st.bar_chart(leads["Status"].value_counts())

    with col7:
        st.subheader("Categorias em estoque")
        if not produtos.empty:
            st.bar_chart(produtos["Categoria"].value_counts())

    st.info("💡 Para prospecção, o cliente verá somente produto, categoria, preço de venda e disponibilidade. Custo e fornecedor ficam escondidos.")


# ==========================================================
# LEADS
# ==========================================================

if menu == "Leads":
    st.subheader("🔎 Busca de Clientes")
    st.caption("Pesquisa simplificada mostrando somente nome, telefone e estado.")

    leads = carregar_base_crm()

    if leads.empty:
        st.warning("Nenhum lead encontrado ainda.")
    else:
        for coluna in ["Nome", "Telefone", "UF"]:
            if coluna not in leads.columns:
                leads[coluna] = ""

        col1, col2 = st.columns([2, 1])

        with col1:
            busca = st.text_input(
                "Buscar por nome ou telefone",
                placeholder="Digite o nome ou o telefone do cliente"
            )

        estados = sorted([
            uf for uf in leads["UF"].fillna("").astype(str).str.strip().unique().tolist()
            if uf
        ])

        with col2:
            estado = st.selectbox("Estado", ["Todos"] + estados)

        resultado = leads.copy()

        if busca:
            termo = busca.strip().lower()
            nome_texto = resultado["Nome"].fillna("").astype(str).str.lower()
            telefone_texto = resultado["Telefone"].fillna("").astype(str).str.lower()
            resultado = resultado[
                nome_texto.str.contains(termo, na=False, regex=False)
                | telefone_texto.str.contains(termo, na=False, regex=False)
            ]

        if estado != "Todos":
            resultado = resultado[
                resultado["UF"].fillna("").astype(str).str.strip() == estado
            ]

        resultado_exibicao = resultado[["Nome", "Telefone", "UF"]].copy()
        resultado_exibicao = resultado_exibicao.rename(columns={"UF": "Estado"})

        st.success(f"💖 {len(resultado_exibicao)} clientes encontrados")
        st.dataframe(
            resultado_exibicao,
            use_container_width=True,
            height=520,
            hide_index=True
        )

        csv = resultado_exibicao.to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            "⬇️ Baixar resultado da busca",
            csv,
            "clientes_nome_telefone_estado.csv",
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
                status_atual = leads.loc[idx, "Status"] if leads.loc[idx, "Status"] in STATUS_OPCOES else "Novo"

                status = col1.selectbox(
                    "Status",
                    STATUS_OPCOES,
                    index=STATUS_OPCOES.index(status_atual)
                )
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
    produtos = carregar_produtos()

    if leads.empty:
        st.warning("Nenhum lead cadastrado.")
    else:
        filtrado = aplicar_filtros(leads)

        if filtrado.empty:
            st.warning("Nenhum lead encontrado nesse filtro.")
        else:
            categorias = ["Todas"]
            if not produtos.empty:
                categorias += sorted([c for c in produtos["Categoria"].dropna().unique().tolist() if limpar_texto(c)])

            col1, col2 = st.columns(2)

            with col1:
                idx = st.selectbox(
                    "Escolha o lead",
                    filtrado.index.tolist(),
                    format_func=lambda i: f'{leads.loc[i, "Nome"]} | {leads.loc[i, "Telefone"]} | {leads.loc[i, "Email"]}'
                )

            with col2:
                categoria_msg = st.selectbox("Categoria para sugerir na mensagem", categorias)

            row = leads.loc[idx]

            col3, col4 = st.columns(2)
            with col3:
                tom = st.radio("Modelo de WhatsApp", ["acolhedora", "curta"], horizontal=True)
            with col4:
                incluir_produtos = st.checkbox("Incluir produtos com preço de venda", value=True)

            msg_wpp = gerar_msg_whatsapp(
                row,
                produtos=produtos,
                categoria=categoria_msg,
                tom=tom,
                incluir_produtos=incluir_produtos
            )

            msg_email = gerar_msg_email(
                row,
                produtos=produtos,
                categoria=categoria_msg,
                incluir_produtos=incluir_produtos
            )

            st.markdown("### WhatsApp")
            msg_wpp_editada = st.text_area("Mensagem para WhatsApp", msg_wpp, height=300)

            link = gerar_link_whatsapp(row.get("Telefone", ""), msg_wpp_editada)
            st.markdown(f"[🚀 Abrir WhatsApp do lead]({link})")

            st.markdown("### E-mail")
            assunto = "Conheça a LuhVee Stores ❤️"
            assunto_editado = st.text_input("Assunto", assunto)
            msg_email_editada = st.text_area("Mensagem para e-mail", msg_email, height=360)

            email = limpar_texto(row.get("Email", ""))

            st.markdown("### Dados para copiar")
            st.text_input("E-mail do lead", email if email else "Sem e-mail cadastrado")
            st.text_input("Assunto para copiar", assunto_editado)
            st.text_area("Mensagem para copiar e colar no e-mail", msg_email_editada, height=260, key="email_para_copiar")

            if email:
                mailto = f"mailto:{email}?subject={urllib.parse.quote(assunto_editado)}&body={urllib.parse.quote(msg_email_editada)}"

                gmail_url = (
                    "https://mail.google.com/mail/?view=cm&fs=1"
                    f"&to={urllib.parse.quote(email)}"
                    f"&su={urllib.parse.quote(assunto_editado)}"
                    f"&body={urllib.parse.quote(msg_email_editada)}"
                )

                col_email1, col_email2 = st.columns(2)

                with col_email1:
                    st.markdown(f"[📧 Abrir no aplicativo de e-mail]({mailto})")

                with col_email2:
                    st.markdown(f"[📩 Abrir no Gmail]({gmail_url})")

                st.info("Se o botão do aplicativo de e-mail não abrir, use o botão 'Abrir no Gmail' ou copie os dados acima.")
            else:
                st.info("Esse lead não tem e-mail cadastrado.")

            st.warning("⚠️ Evite disparo em massa. Personalize a abordagem e respeite pedidos de remoção da lista.")


# ==========================================================
# CATÁLOGO
# ==========================================================

if menu == "Catálogo":
    st.subheader("🛍️ Catálogo à pronta entrega")

    produtos = carregar_produtos()

    if produtos.empty:
        st.warning("Não encontrei estoque_base.xlsx nem produtos.csv na pasta do sistema.")
    else:
        categorias = ["Todas"] + sorted([c for c in produtos["Categoria"].dropna().unique().tolist() if limpar_texto(c)])
        categoria = st.selectbox("Filtrar categoria", categorias)

        busca = st.text_input("Buscar produto")

        visivel = produtos.copy()

        if categoria != "Todas":
            visivel = visivel[visivel["Categoria"] == categoria]

        if busca:
            busca_lower = busca.lower()
            texto = visivel.astype(str).agg(" ".join, axis=1).str.lower()
            visivel = visivel[texto.str.contains(busca_lower, na=False)]

        col1, col2 = st.columns(2)
        col1.metric("Produtos encontrados", len(visivel))
        col2.metric("À pronta entrega", len(visivel[visivel["Estoque"] > 0]))

        st.markdown("### Visão para cliente")
        tabela_cliente = visivel[["Produto", "Categoria", "Preço Venda", "Estoque", "Disponível"]].copy()
        tabela_cliente["Preço Venda"] = tabela_cliente["Preço Venda"].apply(moeda)
        tabela_cliente["Estoque"] = tabela_cliente["Estoque"].astype(int, errors="ignore")
        st.dataframe(tabela_cliente, use_container_width=True, height=420)

        st.download_button(
            "⬇️ Baixar Catálogo para Cliente",
            tabela_cliente.to_csv(index=False).encode("utf-8-sig"),
            "catalogo_cliente_luhvee.csv",
            "text/csv"
        )

        st.markdown("---")
        st.markdown("### Mensagem de catálogo por WhatsApp")

        categoria_msg = st.selectbox("Categoria da mensagem", categorias, key="categoria_catalogo_msg")
        mensagem_catalogo = f"""💖 Catálogo LuhVee Stores ❤️

Produtos à pronta entrega selecionados com carinho:

{catalogo_para_cliente(produtos, categoria=categoria_msg, limite=15)}

Me chama para reservar o seu ou pedir mais fotos ✨"""

        msg_editada = st.text_area("Mensagem pronta", mensagem_catalogo, height=320)
        link = f"https://wa.me/{WHATSAPP_LOJA}?text={urllib.parse.quote(msg_editada)}"
        st.markdown(f"[🚀 Enviar pelo WhatsApp da loja]({link})")

        with st.expander("Controle interno"):
            st.info("Aqui você pode visualizar dados internos. Eles não aparecem nas mensagens para clientes.")
            st.dataframe(produtos, use_container_width=True, height=300)


# ==========================================================
# CAMPANHAS
# ==========================================================

if menu == "Campanhas":
    st.subheader("📢 Campanhas prontas")

    produtos = carregar_produtos()
    categorias = ["Todas"]
    if not produtos.empty:
        categorias += sorted([c for c in produtos["Categoria"].dropna().unique().tolist() if limpar_texto(c)])

    categoria_campanha = st.selectbox("Categoria para campanha", categorias)

    mensagem1 = f"""💖 Olá! Tudo bem?

A LuhVee Stores preparou novidades especiais para quem ama se cuidar e presentear ✨

Temos produtos à pronta entrega selecionados com carinho.

{catalogo_para_cliente(produtos, categoria=categoria_campanha, limite=8)}

Quer receber fotos e mais detalhes?"""

    mensagem2 = f"""✨ Novidades LuhVee Stores ❤️

Produtos para autocuidado, presentes e aquele mimo especial do dia a dia.

Opções disponíveis hoje:

{catalogo_para_cliente(produtos, categoria=categoria_campanha, limite=8)}

Me chama que eu te envio mais informações 💖"""

    mensagem3 = """💖 Presenteie com carinho!

A LuhVee Stores monta kits lindos e personalizados para várias ocasiões 🎁

Temos perfumes árabes originais, hidratantes, body splash, cosméticos e diversos produtos especiais.

Quer ver as opções de hoje?"""

    mensagem_lgpd = """Olá! Sou da LuhVee Stores ❤️

Estou entrando em contato para apresentar nossos produtos de autocuidado, presentes e achadinhos especiais.

Caso não queira receber novas mensagens, é só me avisar que removo seu contato da nossa lista. 💖"""

    st.text_area("💬 Campanha 1 - Apresentação com produtos", mensagem1, height=260)
    st.text_area("💬 Campanha 2 - Novidades com produtos", mensagem2, height=260)
    st.text_area("💬 Campanha 3 - Kits e presentes", mensagem3, height=180)
    st.text_area("💬 Rodapé opcional LGPD", mensagem_lgpd, height=150)


# ==========================================================
# WHATSAPP
# ==========================================================

if menu == "WhatsApp":
    st.subheader("💬 Central WhatsApp")

    produtos = carregar_produtos()
    categorias = ["Todas"]
    if not produtos.empty:
        categorias += sorted([c for c in produtos["Categoria"].dropna().unique().tolist() if limpar_texto(c)])

    categoria = st.selectbox("Categoria para enviar", categorias)

    mensagem_padrao = f"""💖 Olá! Conheça as novidades da LuhVee Stores ❤️

Produtos à pronta entrega:

{catalogo_para_cliente(produtos, categoria=categoria, limite=12)}

Me chama para reservar o seu ✨"""

    mensagem = st.text_area("Mensagem", mensagem_padrao, height=320)

    link = f"https://wa.me/{WHATSAPP_LOJA}?text={urllib.parse.quote(mensagem)}"
    st.markdown(f"[🚀 Abrir WhatsApp da loja]({link})")

    st.info("💡 Use mensagens menores, personalizadas e evite disparos em massa para não bloquear o WhatsApp.")
