import os
import xmlrpc.client
import traceback
import pandas as pd
import streamlit as st
import json
import time
import csv
import ast
import sqlite3
from dotenv import load_dotenv

# Carregar variáveis do .env
load_dotenv()
path_down = os.getenv("CAMINHO_DOWNLOAD") or "."

# Garante que a pasta de download exista
os.makedirs(path_down, exist_ok=True)

# Caminho do banco de dados SQLite
DB_PATH = os.path.join(path_down, "filtros_odoo.db")

# ---------------- Banco de Filtros (SQLite) ---------------- #
def inicializar_banco():
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS filtros_odoo (
            nome TEXT PRIMARY KEY,
            domain TEXT,
            fields TEXT
        )
    ''')
    conn.commit()
    conn.close()

def carregar_filtros_salvos():
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute('SELECT nome, domain, fields FROM filtros_odoo')
    rows = cursor.fetchall()
    conn.close()
    return {
        nome: {"domain": json.loads(domain), "fields": json.loads(fields)}
        for nome, domain, fields in rows
    }

def salvar_filtro(nome_filtro, domain, fields):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute('''
        INSERT OR REPLACE INTO filtros_odoo (nome, domain, fields)
        VALUES (?, ?, ?)
    ''', (nome_filtro, json.dumps(domain), json.dumps(fields)))
    conn.commit()
    conn.close()

def excluir_filtro(nome_filtro):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute('DELETE FROM filtros_odoo WHERE nome = ?', (nome_filtro,))
    conn.commit()
    conn.close()

# Inicializar banco
inicializar_banco()

# ---------------- Funções auxiliares ---------------- #
def corrigir_entrada_json(texto):
    try:
        return json.loads(texto)
    except json.JSONDecodeError:
        try:
            texto_corrigido = texto.replace("'", '"')
            return json.loads(texto_corrigido)
        except:
            try:
                return ast.literal_eval(texto)
            except:
                return None

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

def buscar_movimentacoes(uid, models, db, senha, modelo, domain, fields):
    try:
        offset = 0
        limit = 500
        todos_registros = []
        progresso = st.empty()

        while True:
            ids = models.execute_kw(
                db, uid, senha,
                modelo, 'search',
                [domain], {'offset': offset, 'limit': limit}
            )
            if not ids:
                break

            lote = models.execute_kw(
                db, uid, senha,
                modelo, 'read', [ids], {'fields': fields}
            )
            todos_registros.extend(lote)
            offset += limit

            progresso.info(f"➡️ {len(todos_registros)} registros coletados até agora...")
            time.sleep(0.1)

        progresso.success(f"✅ Total de {len(todos_registros)} registros coletados.")
        return todos_registros

    except Exception as e:
        st.error("Erro ao buscar movimentações:")
        st.exception(e)
        return []

def normalizar_registros(registros, models, db, uid, senha):
    campos_partner = ['parte_contraria_ids', 'parte_representada_ids', 'advogado_adverso_ids']
    todos_partner_ids = set()
    dossie_ids_para_buscar = []

    for registro in registros:
        for campo in campos_partner:
            valor = registro.get(campo, [])
            if isinstance(valor, list) and all(isinstance(v, int) for v in valor) and valor:
                primeiro_id = valor[0]
                todos_partner_ids.add(primeiro_id)
                registro[campo] = primeiro_id

        dossie_id = registro.get("id")
        if dossie_id:
            dossie_ids_para_buscar.append(dossie_id)

    partner_id_to_name = {}
    if todos_partner_ids:
        partner_nomes = models.execute_kw(
            db, uid, senha,
            'res.partner', 'read',
            [list(todos_partner_ids)], {'fields': ['name']}
        )
        partner_id_to_name = {p['id']: p['name'] for p in partner_nomes}

    dossie_id_to_relacionados = {}
    for dossie_id in dossie_ids_para_buscar:
        relacionados = models.execute_kw(
            db, uid, senha,
            'dossie.processo.relacionado', 'search_read',
            [[('dossie_id', '=', dossie_id)]],
            {'fields': ['name'], 'limit': 1}
        )
        dossie_id_to_relacionados[dossie_id] = relacionados[0]['name'] if relacionados else ""

    for registro in registros:
        for chave, valor in registro.items():
            if isinstance(valor, list) and len(valor) == 2 and isinstance(valor[0], int):
                registro[chave] = valor[1]
            elif isinstance(valor, list) and all(isinstance(v, list) and len(v) == 2 for v in valor):
                registro[chave] = ", ".join(v[1] for v in valor)
            elif isinstance(valor, int) and chave in campos_partner:
                registro[chave] = partner_id_to_name.get(valor, str(valor))
            elif isinstance(valor, list) and all(isinstance(v, int) for v in valor):
                registro[chave] = ", ".join(str(v) for v in valor)

        registro['processos_relacionados'] = dossie_id_to_relacionados.get(registro.get("id"), "")

    return registros

def salvar_excel(registros, models, db, uid, senha):
    registros = normalizar_registros(registros, models, db, uid, senha)
    df = pd.DataFrame(registros)
    excel_path = "Extracao.xlsx"
    df.to_excel(excel_path, index=False)
    return excel_path, df

# ---------------- Estado Inicial ---------------- #
if "domain_input" not in st.session_state:
    st.session_state["domain_input"] = '[["estado_cliente", "=", "a"]]'
if "fields_input" not in st.session_state:
    st.session_state["fields_input"] = '["dossie_id", "processo", "fase_id"]'

# --- Carregar filtros salvos e aplicar se necessário --- #
filtros_disponiveis = carregar_filtros_salvos()

if st.session_state.get("aplicar_filtro", False):
    filtro_aplicado = st.session_state.get("filtro_selecionado")
    if filtro_aplicado in filtros_disponiveis:
        st.session_state["domain_input"] = filtros_disponiveis[filtro_aplicado]["domain"]
        st.session_state["fields_input"] = filtros_disponiveis[filtro_aplicado]["fields"]
    st.session_state["aplicar_filtro"] = False

# ---------------- Layout ---------------- #
st.set_page_config(page_title="Exportador Personalizado Odoo", layout="wide")
st.title("🔐 Exportador Personalizado Odoo")

with st.form("form_config"):
    st.subheader("🔧 Configurações de Conexão")
    col1, col2 = st.columns(2)
    with col1:
        url = st.text_input("URL do Odoo", value="https://mmp.intelligenti.com.br")
        db = st.text_input("Banco de Dados", value="mmp.intelligenti.com.br")
    with col2:
        usuario = st.text_input("Usuário")
        senha = st.text_input("Senha", type="password")

    st.subheader("📄 Modelo a Consultar")
    modelo_input = st.text_input("Modelo (ex: dossie.dossie)", value="dossie.dossie")

    with st.expander("📌 Parâmetros da Consulta"):
        domain_input = st.text_area("Filtro", value=st.session_state["domain_input"], key="domain_input")
        fields_input = st.text_area("Campos", value=st.session_state["fields_input"], key="fields_input")

    st.subheader("💾 Gerenciar Filtros Salvos")
    nome_filtro = st.text_input("Nome do Filtro para Salvar")

    col3, col4 = st.columns([1, 1])
    with col3:
        if st.form_submit_button("🔄 Conectar e Buscar Dados"):
            processar = True
        else:
            processar = False
    with col4:
        if nome_filtro and st.form_submit_button("💾 Salvar Filtro Atual"):
            salvar_filtro(nome_filtro, domain_input, fields_input)
            st.success(f"Filtro '{nome_filtro}' salvo com sucesso!")

    if filtros_disponiveis:
        filtro_selecionado = st.selectbox("📂 Carregar Filtro Salvo", list(filtros_disponiveis.keys()))
        col_del, col_apl = st.columns([1, 1])

        with col_del:
            if st.form_submit_button("🗑️ Excluir Filtro Selecionado"):
                excluir_filtro(filtro_selecionado)
                st.success(f"Filtro '{filtro_selecionado}' excluído com sucesso!")
                st.rerun()

        with col_apl:
            if st.form_submit_button("📌 Aplicar Filtro"):
                st.session_state["filtro_selecionado"] = filtro_selecionado
                st.session_state["aplicar_filtro"] = True
                st.rerun()

# ---------------- Execução ---------------- #
if "processar" not in locals():
    processar = False

if processar:
    domain = corrigir_entrada_json(st.session_state["domain_input"])
    fields = corrigir_entrada_json(st.session_state["fields_input"])

    if domain is None or fields is None:
        st.error("❌ Erro ao interpretar domain ou fields. Verifique se estão em formato JSON.")
        st.stop()

    with st.spinner("🔐 Conectando ao Odoo..."):
        uid, models = logar_no_odoo(url, db, usuario, senha)

    if uid:
        with st.spinner("🔍 Buscando Casos..."):
            registros = buscar_movimentacoes(uid, models, db, senha, modelo_input, domain, fields)

        if registros:
            caminho_excel, df = salvar_excel(registros, models, db, uid, senha)
            st.success(f"✅ {len(df)} registros exportados com sucesso!")
            st.markdown(f"### 📊 Visualização dos {len(df)} primeiros registros")
            st.dataframe(df.head(100))

            with open(caminho_excel, "rb") as f:
                st.download_button("📥 Baixar Excel", f, file_name="Extracao.xlsx", type="primary")
        else:
            st.warning("⚠️ Nenhum registro encontrado.")
