import os
import xmlrpc.client
import traceback
import pandas as pd
import streamlit as st
import json
from dotenv import load_dotenv

# Carrega apenas o caminho de download do .env
load_dotenv()
path_down = os.getenv("CAMINHO_DOWNLOAD") or "."

# ------------------------------
# Funções auxiliares
# ------------------------------

def logar_no_odoo(url, db, usuario, senha):
    try:
        common = xmlrpc.client.ServerProxy(f"{url}/xmlrpc/2/common")
        models = xmlrpc.client.ServerProxy(f"{url}/xmlrpc/2/object")
        uid = common.authenticate(db, usuario, senha, {})
        if not uid:
            raise ValueError("Falha na autenticação Odoo")
        return uid, models
    except Exception as e:
        st.error(f"Erro ao autenticar no Odoo: {str(e)}")
        return None, None

def buscar_movimentacoes(uid, models, db, senha, domain, fields):
    try:
        offset = 0
        limit = 1000
        todos_registros = []
        progresso = st.empty()  # espaço dinâmico para exibir progresso

        while True:
            lote = models.execute_kw(
                db, uid, senha,
                'dossie.dossie', 'search_read',
                [domain], {'fields': fields, 'offset': offset, 'limit': limit}
            )
            if not lote:
                break

            todos_registros.extend(lote)
            offset += limit

            progresso.info(f"➡️ {len(todos_registros)} registros coletados até agora...")

        progresso.success(f"✅ Total de {len(todos_registros)} registros coletados.")
        return todos_registros

    except Exception as e:
        st.error("Erro ao buscar movimentações:")
        st.exception(e)
        return []


def normalizar_registros(registros):
    """Converte campos many2one e many2many para strings legíveis."""
    for registro in registros:
        for chave, valor in registro.items():
            # Trata Many2one: [id, "nome"]
            if isinstance(valor, list) and len(valor) == 2 and isinstance(valor[0], int) and isinstance(valor[1], str):
                registro[chave] = valor[1]

            # Trata Many2many: [[id, "nome"], [id, "nome2"], ...]
            elif isinstance(valor, list) and all(isinstance(v, list) and len(v) == 2 for v in valor):
                nomes = [v[1] for v in valor if isinstance(v[1], str)]
                registro[chave] = ", ".join(nomes)

    return registros

def get_download_folder():
    if os.name == 'nt':  # Windows
        download_folder = os.path.join(os.environ['USERPROFILE'], 'Downloads')
    else:  # Linux/Mac
        download_folder = os.path.join(os.environ['HOME'], 'Downloads')
    return download_folder

def salvar_excel(registros):
    registros = normalizar_registros(registros)
    df = pd.DataFrame(registros)
    download_folder = get_download_folder()
    excel_path = os.path.join(download_folder, "Extracao.xlsx")
    df.to_excel(excel_path, index=False)
    return excel_path, df



# ------------------------------
# Streamlit Interface
# ------------------------------

st.set_page_config(page_title="Exportador de Dossiês Personalizado", layout="wide")
st.title("🔐 Exportador de Dossiês -")

with st.form("form_config"):
    st.subheader("🔧 Configurações de Conexão")
    url = st.text_input("URL do Odoo", value="https://mmp.intelligenti.com.br")
    db = st.text_input("Banco de Dados", value="mmp.intelligenti.com.br")
    usuario = st.text_input("Usuário", placeholder="Digite seu login do Odoo")
    senha = st.text_input("Senha", type="password")

    st.subheader("📌 Parâmetros da Consulta")
    domain_input = st.text_area("Filtro", value='[["estado_cliente", "=", "a"]]')
    fields_input = st.text_area("Campos", value='["dossie_id", "processo", "fase_id"]')

    submitted = st.form_submit_button("🔄 Conectar e Buscar Dados")

# Só executa a partir daqui se clicar no botão
if submitted:
    try:
        domain = json.loads(domain_input)
        fields = json.loads(fields_input)
    except Exception as e:
        st.error("❌ Erro ao interpretar domain ou fields. Verifique se estão em formato JSON.")
        st.stop()

    with st.spinner("🔐 Conectando ao Odoo..."):
        uid, models = logar_no_odoo(url, db, usuario, senha)

    if uid:
        with st.spinner("🔍 Buscando Casos..."):
            registros = buscar_movimentacoes(uid, models, db, senha, domain, fields)

        if registros:
            caminho_excel, df = salvar_excel(registros)
            st.success(f"✅ {len(df)} registros exportados com sucesso!")
            st.dataframe(df.head())

            with open(caminho_excel, "rb") as f:
                st.download_button("📥 Baixar Excel", f, file_name="Karol_dossie.xlsx")
        else:
            st.warning("⚠️ Nenhum registro encontrado.")
