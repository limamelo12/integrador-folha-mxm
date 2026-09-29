import streamlit as st
import pandas as pd
import os
import re
import shutil
import tempfile
import urllib.request
import json
from datetime import datetime, timezone, timedelta
import importlib

import pdf_parser
import accounting_engine

try:
    importlib.reload(pdf_parser)
    importlib.reload(accounting_engine)
except Exception:
    pass

from pdf_parser import parse_payroll_pdf
from accounting_engine import load_mapping, save_mapping, generate_entries, write_to_excel_template

@st.cache_data(ttl=120)
def get_last_github_update():
    """Fetches the latest commit timestamp from GitHub, with local fallback."""
    # 1. GitHub API
    try:
        repo = 'limamelo12/integrador-folha-mxm'
        url = f'https://api.github.com/repos/{repo}/commits?per_page=1'
        req = urllib.request.Request(url, headers={'User-Agent': 'Streamlit-App'})
        with urllib.request.urlopen(req, timeout=3) as resp:
            data = json.loads(resp.read().decode('utf-8'))
            commit_date_str = data[0]['commit']['committer']['date']
            dt = datetime.fromisoformat(commit_date_str.replace('Z', '+00:00'))
            try:
                import zoneinfo
                dt_sp = dt.astimezone(zoneinfo.ZoneInfo('America/Sao_Paulo'))
            except Exception:
                dt_sp = dt.astimezone(timezone(timedelta(hours=-3)))
            return dt_sp.strftime('%d/%m/%Y às %H:%M')
    except Exception:
        pass

    # 2. Local fallback
    try:
        target_files = ['app.py', 'accounting_engine.py', 'pdf_parser.py']
        mtimes = [os.path.getmtime(os.path.join(os.path.dirname(os.path.abspath(__file__)), f)) for f in target_files if os.path.exists(os.path.join(os.path.dirname(os.path.abspath(__file__)), f))]
        if mtimes:
            dt = datetime.fromtimestamp(max(mtimes))
            return dt.strftime('%d/%m/%Y às %H:%M')
    except Exception:
        pass

    return datetime.now().strftime('%d/%m/%Y')

# Page Configuration
st.set_page_config(
    page_title="Integrador Contábil MXM - Folha de Pagamento",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Base Directory of the application
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
local_logo = os.path.join(BASE_DIR, "Logo.png")

# Custom Premium Styling (Monte Carlo Branding)
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
    
    /* Force Light Color Scheme Globally */
    :root {
        color-scheme: light !important;
    }

    /* Global fonts and background */
    html, body, [class*="css"], .stText {
        font-family: 'Inter', sans-serif;
    }
    
    .stApp {
        background-color: #fcf8f2 !important; /* Delicate Cream background */
        color: #0b0b0b !important;
    }
    
    /* Ensure main area texts are forced to Onyx Black */
    .stApp p, .stApp span, .stApp label, .stApp li, .stApp div {
        color: #0b0b0b;
    }
    
    /* Force main headings to Onyx Black */
    .stApp h1, .stApp h2, .stApp h3, .stApp h4, .stApp h5, .stApp h6 {
        color: #0b0b0b !important;
    }
    
    /* Sidebar Styling - Light Pérola background to make Black Logo visible */
    [data-testid="stSidebar"] {
        background-color: #eae8e4 !important; /* Pérola */
        border-right: 1px solid #dcdad5;
    }
    
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p, 
    [data-testid="stSidebar"] label,
    [data-testid="stSidebar"] span,
    [data-testid="stSidebar"] li {
        color: #0b0b0b !important; /* Preto Ônix text */
    }
    
    [data-testid="stSidebar"] h1, 
    [data-testid="stSidebar"] h2, 
    [data-testid="stSidebar"] h3 {
        color: #0b0b0b !important; /* Preto Ônix header */
    }
    
    /* Custom headers */
    h1 {
        font-weight: 700;
        color: #0b0b0b !important; /* Preto Ônix */
        margin-bottom: 0.2rem;
    }
    h2, h3, h4 {
        color: #0b0b0b !important;
        font-weight: 600;
    }
    
    /* Metric boxes */
    div[data-testid="stMetricValue"] {
        font-size: 1.8rem;
        font-weight: 700;
        color: #0b0b0b !important;
    }
    div[data-testid="stMetricLabel"] {
        font-size: 0.85rem;
        color: #555555 !important;
        font-weight: 500;
        text-transform: uppercase;
        letter-spacing: 0.05em;
    }
    
    /* File Uploader - Crisp White Box replacing black block */
    [data-testid="stFileUploader"] {
        background-color: #ffffff !important;
        border: 1px solid #dcdad5 !important;
        border-radius: 12px !important;
        padding: 16px !important;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.02) !important;
    }
    [data-testid="stFileUploader"] section {
        background-color: #faf9f6 !important;
        border: 2px dashed #dcdad5 !important;
        border-radius: 8px !important;
    }
    [data-testid="stFileUploader"] section * {
        color: #0b0b0b !important;
    }
    [data-testid="stFileUploader"] section svg {
        fill: #555555 !important;
    }
    [data-testid="stFileUploader"] button {
        background-color: #ffffff !important;
        color: #0b0b0b !important;
        border: 1px solid #dcdad5 !important;
        font-weight: 500 !important;
        box-shadow: 0 1px 2px rgba(0,0,0,0.05) !important;
    }
    [data-testid="stFileUploader"] button:hover {
        background-color: #f4f2ee !important;
        border-color: #c8c6c0 !important;
    }
    
    /* Inputs, Selectboxes, and Textareas - Clean White replacing black inputs */
    div[data-baseweb="input"], 
    div[data-baseweb="base-input"],
    div[data-baseweb="select"] > div {
        background-color: #ffffff !important;
        border: 1px solid #dcdad5 !important;
        border-radius: 8px !important;
        color: #0b0b0b !important;
    }
    div[data-baseweb="input"] input,
    div[data-baseweb="base-input"] input,
    div[data-baseweb="select"] input {
        background-color: #ffffff !important;
        color: #0b0b0b !important;
        -webkit-text-fill-color: #0b0b0b !important;
    }
    div[data-baseweb="select"] * {
        color: #0b0b0b !important;
    }
    div[data-baseweb="select"] svg {
        fill: #0b0b0b !important;
    }

    /* Dataframe and Data Editor Tables - Light White Styling */
    [data-testid="stDataFrame"], 
    [data-testid="stDataEditor"],
    div[data-testid="stDataFrame"] > div,
    div[data-testid="stDataEditor"] > div {
        background-color: #ffffff !important;
        border: 1px solid #e2ded7 !important;
        border-radius: 10px !important;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.02) !important;
    }
    
    /* Card design - Pérola/White hybrid */
    .custom-card {
        background: #ffffff !important;
        color: #0b0b0b !important;
        border-radius: 12px;
        padding: 24px;
        box-shadow: 0 4px 6px -1px rgb(0 0 0 / 0.02), 0 2px 4px -2px rgb(0 0 0 / 0.02);
        border: 1px solid #eae8e4 !important; /* Pérola border */
        margin-bottom: 20px;
    }
    
    .custom-card p, .custom-card td, .custom-card th, .custom-card h4 {
        color: #0b0b0b !important;
    }
    
    /* Expander - Clean White */
    [data-testid="stExpander"] {
        background-color: #ffffff !important;
        border: 1px solid #e2ded7 !important;
        border-radius: 10px !important;
        box-shadow: 0 2px 4px rgba(0,0,0,0.02) !important;
    }
    [data-testid="stExpander"] details {
        background-color: #ffffff !important;
    }
    [data-testid="stExpander"] summary {
        color: #0b0b0b !important;
    }
    [data-testid="stExpander"] summary:hover {
        color: #0b0b0b !important;
    }
    [data-testid="stExpander"] summary svg {
        fill: #0b0b0b !important;
    }

    .status-badge {
        display: inline-block;
        padding: 6px 12px;
        border-radius: 9999px;
        font-size: 0.85rem;
        font-weight: 600;
        text-align: center;
    }
    .status-balanced {
        background-color: #dcfce7 !important;
        color: #166534 !important;
        border: 1px solid #bbf7d0 !important;
    }
    .status-unbalanced {
        background-color: #fee2e2 !important;
        color: #991b1b !important;
        border: 1px solid #fecaca !important;
    }
    
    /* Premium Diamante Amarelo Buttons */
    .stButton>button {
        background: #ffc220 !important; /* Diamante Amarelo background */
        color: #0b0b0b !important; /* Preto Ônix text */
        font-weight: 600 !important;
        border: 1px solid #e0ab1c !important;
        border-radius: 8px !important;
        padding: 0.5rem 1.4rem !important;
        transition: all 0.2s ease !important;
        box-shadow: 0 2px 4px rgba(255, 194, 32, 0.15) !important;
    }
    .stButton>button:hover {
        background: #e0ab1c !important; /* slightly darker gold */
        transform: translateY(-1px);
        box-shadow: 0 4px 12px rgba(255, 194, 32, 0.3) !important;
        color: #0b0b0b !important;
    }
    .stButton>button * {
        color: #0b0b0b !important;
    }
    
    /* Download Buttons */
    .stDownloadButton>button {
        background: #ffc220 !important;
        color: #0b0b0b !important;
        font-weight: 600 !important;
        border: 1px solid #e0ab1c !important;
        border-radius: 8px !important;
        padding: 0.6rem 1.6rem !important;
    }
    .stDownloadButton>button:hover {
        background: #e0ab1c !important;
        color: #0b0b0b !important;
    }
    .stDownloadButton>button * {
        color: #0b0b0b !important;
    }

    /* Radio buttons and checkboxes */
    div[data-testid="stRadio"] label, div[data-testid="stCheckbox"] label {
        color: #0b0b0b !important;
    }

    /* ========================================================= */
    /* Executive SaaS Navigation Cards in Sidebar               */
    /* ========================================================= */
    section[data-testid="stSidebar"] div[data-testid="stRadio"] > label {
        display: none !important; /* Hide widget label */
    }
    section[data-testid="stSidebar"] div[data-testid="stRadio"] div[role="radiogroup"] {
        display: flex !important;
        flex-direction: column !important;
        gap: 8px !important;
        padding-top: 2px !important;
    }
    section[data-testid="stSidebar"] div[data-testid="stRadio"] div[role="radiogroup"] label {
        background-color: #ffffff !important;
        border: 1px solid #e2ded7 !important;
        border-radius: 10px !important;
        padding: 12px 14px !important;
        cursor: pointer !important;
        transition: all 0.2s cubic-bezier(0.16, 1, 0.3, 1) !important;
        display: flex !important;
        align-items: center !important;
        width: 100% !important;
        box-shadow: 0 1px 3px rgba(0, 0, 0, 0.02) !important;
        margin: 0 !important;
    }
    /* Suppress the default radio circle completely */
    section[data-testid="stSidebar"] div[data-testid="stRadio"] div[role="radiogroup"] label > div:first-child {
        display: none !important;
    }
    /* Typography inside inactive items */
    section[data-testid="stSidebar"] div[data-testid="stRadio"] div[role="radiogroup"] label p {
        font-size: 0.90rem !important;
        font-weight: 500 !important;
        color: #334155 !important;
        margin: 0 !important;
        letter-spacing: -0.01em !important;
        transition: color 0.15s ease !important;
    }
    /* Hover state for inactive items */
    section[data-testid="stSidebar"] div[data-testid="stRadio"] div[role="radiogroup"] label:hover {
        background-color: #f8fafc !important;
        border-color: #0b0b0b !important;
        transform: translateX(4px) !important;
        box-shadow: 0 4px 10px rgba(0, 0, 0, 0.05) !important;
    }
    section[data-testid="stSidebar"] div[data-testid="stRadio"] div[role="radiogroup"] label:hover p {
        color: #0b0b0b !important;
    }
    /* Active selected SaaS Item */
    section[data-testid="stSidebar"] div[data-testid="stRadio"] div[role="radiogroup"] label:has(input:checked),
    section[data-testid="stSidebar"] div[data-testid="stRadio"] div[role="radiogroup"] label[data-checked="true"] {
        background-color: #0b0b0b !important;
        border-color: #0b0b0b !important;
        border-left: 5px solid #ffc220 !important;
        box-shadow: 0 4px 14px rgba(11, 11, 11, 0.22) !important;
        transform: translateX(2px) !important;
    }
    section[data-testid="stSidebar"] div[data-testid="stRadio"] div[role="radiogroup"] label:has(input:checked) p,
    section[data-testid="stSidebar"] div[data-testid="stRadio"] div[role="radiogroup"] label[data-checked="true"] p {
        color: #ffffff !important;
        font-weight: 600 !important;
    }

    /* Footer */
    .footer {
        text-align: center;
        margin-top: 4rem;
        padding: 20px;
        color: #94a3b8 !important;
        font-size: 0.85rem;
        border-top: 1px solid #eae8e4 !important;
    }

    /* Hub Module Cards & Styling */
    .module-card {
        background-color: #ffffff !important;
        border: 1px solid #e2ded7 !important;
        border-radius: 14px;
        padding: 24px;
        margin-bottom: 20px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.02), 0 2px 4px -2px rgba(0, 0, 0, 0.02);
        transition: all 0.2s ease-in-out;
    }
    .module-card:hover {
        border-color: #0b0b0b !important;
        box-shadow: 0 8px 20px rgba(0, 0, 0, 0.06);
        transform: translateY(-2px);
    }
    .badge-active {
        background-color: #ecfdf5;
        color: #047857;
        font-weight: 700;
        font-size: 0.72rem;
        padding: 4px 10px;
        border-radius: 9999px;
        border: 1px solid #a7f3d0;
        letter-spacing: 0.04em;
        text-transform: uppercase;
        display: inline-block;
    }
    .badge-soon {
        background-color: #f8fafc;
        color: #64748b;
        font-weight: 600;
        font-size: 0.72rem;
        padding: 4px 10px;
        border-radius: 9999px;
        border: 1px solid #e2e8f0;
        letter-spacing: 0.04em;
        text-transform: uppercase;
        display: inline-block;
    }
    .tag-chip {
        background-color: #f5f4f0;
        color: #475569;
        font-size: 0.75rem;
        padding: 3px 9px;
        border-radius: 6px;
        display: inline-block;
        margin-right: 6px;
        margin-bottom: 6px;
        font-weight: 500;
        border: 1px solid #e8e6e1;
    }
</style>

""", unsafe_allow_html=True)

# Helper to find template file
TEMPLATE_FILE = os.path.join(BASE_DIR, "Lançamentos Contábeis MXM.xlsx")




# -------------------------------------------------------------
# Module Views & Navigation
# -------------------------------------------------------------

def render_home_page():
    # Top Header Bar
    st.markdown("""
    <div style='display: flex; justify-content: space-between; align-items: center; border-bottom: 1px solid #eae8e4; padding-bottom: 12px; margin-bottom: 25px; font-size: 0.88rem; color: #64748b;'>
        <div>
            <span style='font-weight: 700; color: #0b0b0b; letter-spacing: 0.05em;'>MONTE CARLO</span> 
            &nbsp;|&nbsp; Portal de Automações Contábeis
        </div>
        <div style='display: flex; gap: 15px; align-items: center;'>
            <span style='background: #ecfdf5; color: #047857; font-weight: 600; padding: 2px 10px; border-radius: 9999px; font-size: 0.75rem; border: 1px solid #a7f3d0;'>🟢 ERP MXM Online</span>
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    # Hero Section matching official portal aesthetic
    col_h1, col_h2, col_h3 = st.columns([1, 2.2, 1])
    with col_h2:
        if os.path.exists("Logo.png"):
            st.image("Logo.png", use_container_width=True)
        st.markdown("""
        <div style='text-align: center; margin-top: -10px; margin-bottom: 30px;'>
            <h2 style='font-size: 1.85rem; font-weight: 700; color: #0b0b0b; margin-bottom: 6px;'>
                Portal Contábil Monte Carlo
            </h2>
            <p style='color: #64748b; font-size: 1.0rem;'>
                Central unificada de rotinas contábeis, conciliações e integrações com o ERP MXM.
            </p>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("### 📌 Rotinas & Módulos Disponíveis")
    st.markdown("<p style='color: #64748b; margin-top: -10px; margin-bottom: 20px;'>Selecione uma das automações contábeis abaixo para iniciar:</p>", unsafe_allow_html=True)
    
    # Cards Grid (2 columns layout)
    # Cards Grid (3 columns layout)
    col1, col2, col3 = st.columns(3)
    
    with col1:
        # Card 1: Integrador de Folha (MXM) - Active
        st.markdown("""
        <div class="custom-card" style='border-left: 4px solid #047857; min-height: 310px;'>
            <div style='display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 12px;'>
                <div style='font-size: 1.6rem;'>📊</div>
                <span class='badge-active'>🟢 Operacional</span>
            </div>
            <h4 style='margin-bottom: 6px; font-size: 1.2rem; color: #0b0b0b;'>Integrador de Folha de Pagamento</h4>
            <p style='color: #475569; font-size: 0.88rem; line-height: 1.45; min-height: 65px;'>
                Importação de relatórios Alterdata e geração automática do layout contábil MXM com cruzamento De-Para e conciliação de encargos.
            </p>
            <div style='margin-bottom: 18px;'>
                <span class='tag-chip'>Folha Normal</span>
                <span class='tag-chip'>Rescisão</span>
                <span class='tag-chip'>Férias</span>
                <span class='tag-chip'>Pró-Labore</span>
            </div>
        </div>
        """, unsafe_allow_html=True)
        if st.button("Acessar Integrador de Folha ➔", key="btn_hub_folha", type="primary", use_container_width=True):
            st.session_state.current_page = "folha"
            st.rerun()

    with col2:
        # Card 2: Fechamento Contábil & Auditoria - Operational
        st.markdown("""
        <div class="custom-card" style='border-left: 4px solid #047857; min-height: 310px;'>
            <div style='display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 12px;'>
                <div style='font-size: 1.6rem;'>📋</div>
                <span class='badge-active'>🟢 Operacional</span>
            </div>
            <h4 style='margin-bottom: 6px; font-size: 1.2rem; color: #0b0b0b;'>Fechamento Contábil & Auditoria</h4>
            <p style='color: #475569; font-size: 0.88rem; line-height: 1.45; min-height: 65px;'>
                Auditoria de balancete, detecção automática de contas viradas, apuração de equivalência patrimonial (MEP) e checklist mensal.
            </p>
            <div style='margin-bottom: 18px;'>
                <span class='tag-chip'>Contas Viradas</span>
                <span class='tag-chip'>Equivalência (MEP)</span>
                <span class='tag-chip'>Auditoria Balancete</span>
                <span class='tag-chip'>Checklist Mensal</span>
            </div>
        </div>
        """, unsafe_allow_html=True)
        if st.button("Acessar Fechamento Contábil ➔", key="btn_hub_fechamento", type="primary", use_container_width=True):
            st.session_state.current_page = "fechamento"
            st.rerun()

    with col3:
        # Card 3: Ativo Imobilizado & Depreciação
        st.markdown("""
        <div class="custom-card" style='border-left: 4px solid #cbd5e1; min-height: 310px;'>
            <div style='display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 12px;'>
                <div style='font-size: 1.6rem;'>🏭</div>
                <span class='badge-soon'>🟡 Em Desenvolvimento</span>
            </div>
            <h4 style='margin-bottom: 6px; font-size: 1.2rem; color: #0b0b0b;'>Ativo Imobilizado & Depreciação</h4>
            <p style='color: #475569; font-size: 0.88rem; line-height: 1.45; min-height: 65px;'>
                Controle patrimonial de máquinas da fábrica (3FAB) e benfeitorias em lojas de shopping, com cálculo e geração automática de quotas mensais.
            </p>
            <div style='margin-bottom: 18px;'>
                <span class='tag-chip'>Bens Fabris</span>
                <span class='tag-chip'>Obras e Lojas</span>
                <span class='tag-chip'>Quotas Depreciação</span>
            </div>
        </div>
        """, unsafe_allow_html=True)
        if st.button("Ver Detalhes do Módulo ➔", key="btn_hub_ativo", use_container_width=True):
            st.session_state.current_page = "ativo"
            st.rerun()




def render_fechamento_contabil():
    import io
    
    # Top Bar / Breadcrumb
    col_top_back, col_top_title = st.columns([1.5, 4])
    with col_top_back:
        if st.button("⬅ Hub de Rotinas", key="btn_back_home_fechamento"):
            st.session_state.current_page = "home"
            st.rerun()
    with col_top_title:
        st.markdown("<div style='font-size: 0.88rem; color: #64748b; padding-top: 6px;'><b>Portal Contábil Monte Carlo</b> &nbsp;/&nbsp; <span>Fechamento Contábil & Auditoria</span></div>", unsafe_allow_html=True)

    st.markdown("## 📋 Fechamento Contábil & Auditoria")
    st.markdown("<p style='color: #475569; font-size: 1.02rem; margin-top: -8px;'>Rotinas e ferramentas avançadas para encerramento de período, auditoria de balancete e conciliações societárias.</p>", unsafe_allow_html=True)
    st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)

    tab_viradas, tab_mep, tab_checklist = st.tabs([
        "🔍 Verificação de Contas Viradas",
        "⚖️ Equivalência Patrimonial (MEP)",
        "📋 Checklist de Fechamento"
    ])

    # -------------------------------------------------------------
    # TAB 1: Auditoria de Contas Viradas (Balancete)
    # -------------------------------------------------------------
    with tab_viradas:
        st.markdown("### 🔍 Verificação de Contas Viradas (Saldos Invertidos)")
        st.markdown("Analisa o balancete contábil verificando se contas de Ativo, Passivo, Custos, Despesas e Receitas apresentam saldos contrários à sua natureza contábil, diferenciando contas retificadoras legítimas de inconsistências operacionais.")

        col_cfg1, col_cfg2, col_cfg3 = st.columns([1.5, 1.2, 1.8])
        with col_cfg1:
            emp_auditoria = st.selectbox(
                "Empresa Auditada",
                options=["0001", "0002", "0003"],
                format_func=lambda x: {
                    "0001": "0001 - Monte Carlo Joias (Via Parque)",
                    "0002": "0002 - Jasper Forest Participações (Holding)",
                    "0003": "0003 - MC Indústria de Joias (Fábrica)"
                }.get(x, x),
                key="sel_emp_auditoria"
            )
        with col_cfg2:
            periodo_ref = st.text_input("Período de Referência", value="05/2026", key="txt_periodo_auditoria", help="Mês e ano do balancete contábil analisado.")
        with col_cfg3:
            modo_dados = st.radio(
                "Origem dos Dados do Balancete:",
                options=["🧪 Balancete Demonstrativo Monte Carlo", "📤 Importar Balancete (Excel / CSV)"],
                horizontal=True,
                key="radio_modo_balancete"
            )

        df_raw = None
        if modo_dados.startswith("🧪"):
            # Dados demonstrativos realistas da Monte Carlo
            demo_data = [
                {"conta": "1.1.1.01.0001", "descricao": "Caixa Geral - Lojas", "saldo": 28450.00, "dc": "D"},
                {"conta": "1.1.1.02.0001", "descricao": "Banco Itaú - Conta Movimento", "saldo": 145200.50, "dc": "D"},
                {"conta": "1.1.1.02.0002", "descricao": "Banco Santander - Conta Movimento", "saldo": 15420.00, "dc": "C"},
                {"conta": "1.1.2.01.0001", "descricao": "Clientes a Receber - Cartões de Crédito", "saldo": 680000.00, "dc": "D"},
                {"conta": "1.1.2.01.0005", "descricao": "(-) Perdas Estimadas com Créditos (PCLD)", "saldo": 12500.00, "dc": "C"},
                {"conta": "1.1.3.01.0001", "descricao": "Estoques de Joias Acabadas", "saldo": 1250000.00, "dc": "D"},
                {"conta": "1.2.3.01.0001", "descricao": "Máquinas e Equipamentos da Fábrica", "saldo": 840000.00, "dc": "D"},
                {"conta": "1.2.3.01.0010", "descricao": "(-) Depreciação Acumulada - Máquinas", "saldo": 154000.00, "dc": "C"},
                {"conta": "2.1.1.01.0001", "descricao": "Fornecedores Nacionais de Ouro e Gemas", "saldo": 310000.00, "dc": "C"},
                {"conta": "2.1.1.01.0088", "descricao": "Fornecedor XYZ - Pagamento a Maior", "saldo": 8950.00, "dc": "D"},
                {"conta": "2.1.2.01.0001", "descricao": "Salários e Ordenados a Pagar", "saldo": 98400.00, "dc": "C"},
                {"conta": "2.1.2.01.0005", "descricao": "INSS a Recolher", "saldo": 34200.00, "dc": "C"},
                {"conta": "2.1.2.01.0002", "descricao": "FGTS a Recolher", "saldo": 18250.00, "dc": "C"},
                {"conta": "5.1.1.01.0001", "descricao": "Despesas com Pessoal - Salários", "saldo": 450000.00, "dc": "D"},
                {"conta": "5.1.1.01.0030", "descricao": "Despesas com Viagens e Representação", "saldo": 2400.00, "dc": "C"},
                {"conta": "4.1.1.01.0001", "descricao": "Receita Bruta com Venda de Joias", "saldo": 2150000.00, "dc": "C"},
                {"conta": "4.1.1.02.0001", "descricao": "(-) Devoluções e Cancelamentos de Vendas", "saldo": 24500.00, "dc": "D"},
                {"conta": "4.1.1.02.0005", "descricao": "(-) ICMS e Impostos sobre Vendas", "saldo": 185000.00, "dc": "D"}
            ]
            df_raw = pd.DataFrame(demo_data)
        else:
            file_balancete = st.file_uploader(
                "Carregar Arquivo do Balancete (Excel ou CSV exportado do MXM):",
                type=["xlsx", "xls", "csv"],
                key="uploader_balancete"
            )
            if file_balancete is not None:
                try:
                    if file_balancete.name.lower().endswith(".csv"):
                        df_raw = pd.read_csv(file_balancete, sep=None, engine='python')
                    else:
                        df_raw = pd.read_excel(file_balancete)
                    st.success(f"Arquivo '{file_balancete.name}' carregado com sucesso ({len(df_raw)} linhas).")
                except Exception as e:
                    st.error(f"Erro ao ler arquivo de balancete: {e}")

        if df_raw is not None and not df_raw.empty:
            # Flexible column finder
            def get_col(candidates):
                for c in df_raw.columns:
                    c_clean = str(c).lower().replace(".", "").replace("_", " ").strip()
                    for cand in candidates:
                        if cand in c_clean:
                            return c
                return None

            c_conta = get_col(["conta", "codigo", "classificacao", "cod"])
            c_desc = get_col(["descricao", "nome", "titulo", "desc"])
            c_saldo = get_col(["saldo", "saldo final", "saldo atual", "valor", "atual"])
            c_dc = get_col(["dc", "natureza", "debito credito", "tipo", "indicador"])

            if not c_conta or not c_saldo:
                st.warning("⚠️ Não foi possível identificar automaticamente as colunas de Conta e Saldo. Selecione-as manualmente abaixo:")
                col_sel1, col_sel2, col_sel3 = st.columns(3)
                with col_sel1:
                    c_conta = st.selectbox("Coluna da Conta:", options=list(df_raw.columns), index=0)
                with col_sel2:
                    c_desc = st.selectbox("Coluna da Descrição:", options=list(df_raw.columns), index=min(1, len(df_raw.columns)-1))
                with col_sel3:
                    c_saldo = st.selectbox("Coluna do Saldo:", options=list(df_raw.columns), index=min(2, len(df_raw.columns)-1))

            # Audit processing
            audit_rows = []
            for _, row in df_raw.iterrows():
                conta_val = str(row[c_conta]).strip() if pd.notna(row[c_conta]) else ""
                desc_val = str(row[c_desc]).strip() if c_desc and pd.notna(row[c_desc]) else ""
                saldo_raw = row[c_saldo] if pd.notna(row[c_saldo]) else 0.0
                dc_raw = str(row[c_dc]).strip().upper() if c_dc and pd.notna(row[c_dc]) else ""

                if not conta_val or conta_val.lower() in ("nan", "total", "subtotal"):
                    continue

                # Parse numeric saldo
                try:
                    if isinstance(saldo_raw, str):
                        s_clean = saldo_raw.replace("R$", "").replace(" ", "").replace(".", "").replace(",", ".")
                        val_num = float(s_clean)
                    else:
                        val_num = float(saldo_raw)
                except Exception:
                    val_num = 0.0

                if abs(val_num) < 0.001:
                    continue

                # Determine D or C
                if dc_raw in ("D", "DEB", "DEBITO"):
                    detected_dc = "D"
                elif dc_raw in ("C", "CRED", "CREDITO"):
                    detected_dc = "C"
                else:
                    detected_dc = "C" if val_num < 0 else "D"

                abs_val = abs(val_num)
                clean_digits = re.sub(r"\D", "", conta_val)
                first_digit = clean_digits[0] if clean_digits else ""
                desc_upper = desc_val.upper()

                # Detect redutora
                is_redutora = any(k in desc_upper for k in [
                    "(-)", "DEPRECIA", "AMORTIZA", "PCLD", "PERDAS", "PROVISAO PARA", 
                    "REDUTORA", "PREJUIZO", "ACOES EM TESOURARIA", "DEVOLU", "DESCONTO", "DEDUCAO"
                ])

                expected_dc = "D"
                status = "REGULAR"
                diagnostico = "Saldo regular em conformidade com a natureza da conta."
                grupo_nome = {
                    "1": "Ativo",
                    "2": "Passivo / PL",
                    "3": "Custos",
                    "4": "Receitas",
                    "5": "Despesas"
                }.get(first_digit, "Outros")

                if first_digit == "1":
                    expected_dc = "C" if is_redutora else "D"
                    if detected_dc != expected_dc:
                        status = "INVERTIDA"
                        diagnostico = "Ativo com saldo credor. Indício de conta corrente com saldo negativo (cheque especial) ou adiantamento a reclassificar para o Passivo."
                    elif is_redutora:
                        status = "REDUTORA"
                        diagnostico = "Conta retificadora do Ativo com saldo credor regular."

                elif first_digit == "2":
                    expected_dc = "D" if is_redutora else "C"
                    if detected_dc != expected_dc:
                        status = "INVERTIDA"
                        diagnostico = "Passivo com saldo devedor. Pagamento em duplicidade ou antecipação a fornecedores que deve ser reclassificado para o Ativo Circulante."
                    elif is_redutora:
                        status = "REDUTORA"
                        diagnostico = "Conta retificadora do Patrimônio Líquido com saldo devedor regular."

                elif first_digit in ("3", "5"):
                    expected_dc = "D"
                    if detected_dc != "D":
                        status = "INVERTIDA"
                        diagnostico = "Conta de despesa/custo com saldo credor. Verificar se houve estorno lançado a maior ou crédito indevido em conta de resultado."

                elif first_digit == "4":
                    expected_dc = "D" if is_redutora else "C"
                    if detected_dc != expected_dc:
                        status = "INVERTIDA"
                        diagnostico = "Conta de receita com saldo devedor fora do grupo de deduções. Verificar estorno ou inversão de lançamento."
                    elif is_redutora:
                        status = "REDUTORA"
                        diagnostico = "Dedução da receita bruta com saldo devedor regular."

                audit_rows.append({
                    "Conta": conta_val,
                    "Descrição": desc_val,
                    "Grupo": grupo_nome,
                    "Natureza Esperada": expected_dc,
                    "D/C Atual": detected_dc,
                    "Saldo (R$)": abs_val,
                    "Status": status,
                    "Parecer Técnico / Diagnóstico": diagnostico
                })

            df_audit = pd.DataFrame(audit_rows)

            tot_analisadas = len(df_audit)
            tot_invertidas = (df_audit["Status"] == "INVERTIDA").sum()
            tot_redutoras = (df_audit["Status"] == "REDUTORA").sum()
            tot_regulares = (df_audit["Status"] == "REGULAR").sum()
            val_divergencia = df_audit[df_audit["Status"] == "INVERTIDA"]["Saldo (R$)"].sum()

            # Metric Cards
            m1, m2, m3, m4 = st.columns(4)
            with m1:
                st.metric("Total de Contas", f"{tot_analisadas}")
            with m2:
                st.metric("Contas Regulares", f"{tot_regulares}", delta="✅ Conforme", delta_color="normal")
            with m3:
                st.metric("Contas Viradas (Invertidas)", f"{tot_invertidas}", delta=f"-R$ {val_divergencia:,.2f}" if tot_invertidas > 0 else "Nenhuma", delta_color="inverse")
            with m4:
                st.metric("Contas Redutoras", f"{tot_redutoras}", delta="🟡 Retificadoras", delta_color="off")

            st.markdown("---")

            # Filter options
            col_f1, col_f2 = st.columns([2, 1])
            with col_f1:
                filtro_inconsistencias = st.checkbox("🎯 Exibir somente contas viradas (Inconsistências para correção)", value=(tot_invertidas > 0))
            with col_f2:
                grupos_sel = st.multiselect("Filtrar Grupos:", options=["Ativo", "Passivo / PL", "Despesas", "Receitas", "Custos"], default=["Ativo", "Passivo / PL", "Despesas", "Receitas", "Custos"])

            df_display = df_audit.copy()
            if filtro_inconsistencias:
                df_display = df_display[df_display["Status"] == "INVERTIDA"]
            if grupos_sel:
                df_display = df_display[df_display["Grupo"].isin(grupos_sel)]

            # Styling table with status badges
            def format_status(val):
                if val == "INVERTIDA":
                    return "background-color: #fee2e2; color: #991b1b; font-weight: 700;"
                elif val == "REDUTORA":
                    return "background-color: #fef3c7; color: #92400e; font-weight: 600;"
                else:
                    return "background-color: #ecfdf5; color: #047857; font-weight: 600;"

            st.dataframe(
                df_display.style.applymap(format_status, subset=["Status"]).format({
                    "Saldo (R$)": "R$ {:,.2f}"
                }),
                use_container_width=True,
                hide_index=True
            )

            # Export Excel
            col_exp1, col_exp2 = st.columns([1.5, 3])
            with col_exp1:
                excel_buffer = io.BytesIO()
                with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                    df_audit.to_excel(writer, index=False, sheet_name="Auditoria_Balancete")
                excel_buffer.seek(0)
                st.download_button(
                    "📥 Baixar Relatório de Auditoria (.xlsx)",
                    data=excel_buffer.getvalue(),
                    file_name=f"Auditoria_Balancete_Empresa_{emp_auditoria}_{periodo_ref.replace('/', '_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

    # -------------------------------------------------------------
    # TAB 2: Equivalência Patrimonial (MEP)
    # -------------------------------------------------------------
    with tab_mep:
        st.markdown("### ⚖️ Apuração e Geração de Equivalência Patrimonial (MEP)")
        st.markdown("Cálculo da variação patrimonial nas investidas da Holding (**Empresa 0002 - Jasper Forest Participações Ltda**) e geração do lote contábil compatível com o ERP MXM.")

        st.markdown("""
        <div style='background: #ffffff; border: 1px solid #e2ded7; border-radius: 10px; padding: 15px; margin-bottom: 20px;'>
            <table style='width: 100%; border-collapse: collapse;'>
                <tr>
                    <td style='color: #64748b; font-size: 0.85rem; width: 25%; font-weight: 600;'>EMPRESA INVESTIDORA (HOLDING):</td>
                    <td style='color: #0b0b0b; font-size: 0.95rem; font-weight: 700;'>0002 - JASPER FOREST PARTICIPAÇÕES LTDA</td>
                    <td style='color: #64748b; font-size: 0.85rem; width: 18%; font-weight: 600;'>CENTRO DE CUSTO (DRE):</td>
                    <td style='color: #0b0b0b; font-size: 0.95rem; font-weight: 700;'>2DIR (Diretoria)</td>
                </tr>
            </table>
        </div>
        """, unsafe_allow_html=True)

        col_mep1, col_mep2, col_mep3 = st.columns([1.2, 1.2, 1.2])
        with col_mep1:
            mes_mep = st.text_input("Mês/Ano de Apuração", value="05/2026", key="txt_mes_mep")
        with col_mep2:
            data_mep = st.text_input("Data do Lançamento (DDMMYYYY)", value="31052026", key="txt_data_mep")
        with col_mep3:
            lote_mep = st.text_input("Lote Contábil MXM", value="MEP05", max_chars=6, key="txt_lote_mep")

        st.markdown("#### 🏢 Investidas do Grupo Monte Carlo")

        col_inv1, col_inv2 = st.columns(2)
        with col_inv1:
            st.markdown("""
            <div style='background: #fdfcfb; border: 1px solid #e2ded7; border-radius: 10px; padding: 16px; margin-bottom: 12px; border-left: 4px solid #ffc220;'>
                <b style='font-size: 1.05rem; color: #0b0b0b;'>0001 - Monte Carlo Joias (Via Parque)</b><br>
                <span style='font-size: 0.78rem; color: #64748b;'>Conta Ativo: <b>1220100001</b> | Conta DRE: <b>4110300001</b></span>
            </div>
            """, unsafe_allow_html=True)
            lucro_0001 = st.number_input("Resultado / Lucro Líquido do Mês (R$) - 0001:", value=850000.00, step=10000.0, format="%.2f", key="inp_lucro_0001")
            part_0001 = st.number_input("% de Participação Societária (0001):", value=99.99, step=0.01, format="%.2f", key="inp_part_0001")
            mep_val_0001 = lucro_0001 * (part_0001 / 100.0)
            st.markdown(f"**Resultado de MEP (0001):** `R$ {mep_val_0001:,.2f}`")

        with col_inv2:
            st.markdown("""
            <div style='background: #fdfcfb; border: 1px solid #e2ded7; border-radius: 10px; padding: 16px; margin-bottom: 12px; border-left: 4px solid #ffc220;'>
                <b style='font-size: 1.05rem; color: #0b0b0b;'>0003 - MC Indústria de Joias (Fábrica)</b><br>
                <span style='font-size: 0.78rem; color: #64748b;'>Conta Ativo: <b>1220100002</b> | Conta DRE: <b>4110300001</b></span>
            </div>
            """, unsafe_allow_html=True)
            lucro_0003 = st.number_input("Resultado / Lucro Líquido do Mês (R$) - 0003:", value=320000.00, step=10000.0, format="%.2f", key="inp_lucro_0003")
            part_0003 = st.number_input("% de Participação Societária (0003):", value=99.99, step=0.01, format="%.2f", key="inp_part_0003")
            mep_val_0003 = lucro_0003 * (part_0003 / 100.0)
            st.markdown(f"**Resultado de MEP (0003):** `R$ {mep_val_0003:,.2f}`")

        total_mep_periodo = mep_val_0001 + mep_val_0003
        badge_mep_tipo = "🟢 Receita de MEP no DRE (2DIR)" if total_mep_periodo >= 0 else "🔴 Despesa de MEP no DRE (2DIR)"
        st.markdown(f"""
        <div style='background: #ecfdf5; border: 1px solid #a7f3d0; border-radius: 10px; padding: 16px; margin-top: 15px; margin-bottom: 20px; display: flex; justify-content: space-between; align-items: center;'>
            <div>
                <span style='color: #065f46; font-size: 0.85rem; font-weight: 700; text-transform: uppercase;'>Total de Equivalência Patrimonial Apurado (Holding 0002):</span><br>
                <span style='color: #047857; font-size: 1.6rem; font-weight: 700;'>R$ {total_mep_periodo:,.2f}</span>
            </div>
            <div>
                <span class='badge-active'>{badge_mep_tipo}</span>
            </div>
        </div>
        """, unsafe_allow_html=True)

        # Build MXM Entry Rows
        mep_entries = []
        seq = 1

        # Investida 0001
        if abs(mep_val_0001) >= 0.01:
            if mep_val_0001 > 0:
                mep_entries.append({
                    'empresa': '0002', 'lote': lote_mep, 'data': data_mep, 'documento': lote_mep,
                    'conta': '1220100001', 'cc': None, 'tipo': 'D',
                    'historico': f"VLR REF APURACAO DE EQUIVALENCIA PATRIMONIAL S/ MONTE CARLO JOIAS CONF BALANCETE {mes_mep}".upper()[:200],
                    'valor': mep_val_0001, 'sequencia': seq, 'numero_titulo': ''
                })
                seq += 1
                mep_entries.append({
                    'empresa': '0002', 'lote': lote_mep, 'data': data_mep, 'documento': lote_mep,
                    'conta': '4110300001', 'cc': '2DIR', 'tipo': 'C',
                    'historico': f"VLR REF APURACAO DE EQUIVALENCIA PATRIMONIAL S/ MONTE CARLO JOIAS CONF BALANCETE {mes_mep}".upper()[:200],
                    'valor': mep_val_0001, 'sequencia': seq, 'numero_titulo': ''
                })
                seq += 1
            else:
                mep_entries.append({
                    'empresa': '0002', 'lote': lote_mep, 'data': data_mep, 'documento': lote_mep,
                    'conta': '5110300001', 'cc': '2DIR', 'tipo': 'D',
                    'historico': f"VLR REF PERDA POR EQUIVALENCIA PATRIMONIAL S/ MONTE CARLO JOIAS CONF BALANCETE {mes_mep}".upper()[:200],
                    'valor': abs(mep_val_0001), 'sequencia': seq, 'numero_titulo': ''
                })
                seq += 1
                mep_entries.append({
                    'empresa': '0002', 'lote': lote_mep, 'data': data_mep, 'documento': lote_mep,
                    'conta': '1220100001', 'cc': None, 'tipo': 'C',
                    'historico': f"VLR REF PERDA POR EQUIVALENCIA PATRIMONIAL S/ MONTE CARLO JOIAS CONF BALANCETE {mes_mep}".upper()[:200],
                    'valor': abs(mep_val_0001), 'sequencia': seq, 'numero_titulo': ''
                })
                seq += 1

        # Investida 0003
        if abs(mep_val_0003) >= 0.01:
            if mep_val_0003 > 0:
                mep_entries.append({
                    'empresa': '0002', 'lote': lote_mep, 'data': data_mep, 'documento': lote_mep,
                    'conta': '1220100002', 'cc': None, 'tipo': 'D',
                    'historico': f"VLR REF APURACAO DE EQUIVALENCIA PATRIMONIAL S/ MC INDUSTRIA CONF BALANCETE {mes_mep}".upper()[:200],
                    'valor': mep_val_0003, 'sequencia': seq, 'numero_titulo': ''
                })
                seq += 1
                mep_entries.append({
                    'empresa': '0002', 'lote': lote_mep, 'data': data_mep, 'documento': lote_mep,
                    'conta': '4110300001', 'cc': '2DIR', 'tipo': 'C',
                    'historico': f"VLR REF APURACAO DE EQUIVALENCIA PATRIMONIAL S/ MC INDUSTRIA CONF BALANCETE {mes_mep}".upper()[:200],
                    'valor': mep_val_0003, 'sequencia': seq, 'numero_titulo': ''
                })
                seq += 1
            else:
                mep_entries.append({
                    'empresa': '0002', 'lote': lote_mep, 'data': data_mep, 'documento': lote_mep,
                    'conta': '5110300001', 'cc': '2DIR', 'tipo': 'D',
                    'historico': f"VLR REF PERDA POR EQUIVALENCIA PATRIMONIAL S/ MC INDUSTRIA CONF BALANCETE {mes_mep}".upper()[:200],
                    'valor': abs(mep_val_0003), 'sequencia': seq, 'numero_titulo': ''
                })
                seq += 1
                mep_entries.append({
                    'empresa': '0002', 'lote': lote_mep, 'data': data_mep, 'documento': lote_mep,
                    'conta': '1220100002', 'cc': None, 'tipo': 'C',
                    'historico': f"VLR REF PERDA POR EQUIVALENCIA PATRIMONIAL S/ MC INDUSTRIA CONF BALANCETE {mes_mep}".upper()[:200],
                    'valor': abs(mep_val_0003), 'sequencia': seq, 'numero_titulo': ''
                })
                seq += 1

        st.markdown("#### 📄 Lançamentos Contábeis Gerados para o MXM")
        df_mep_show = pd.DataFrame(mep_entries)
        if not df_mep_show.empty:
            st.dataframe(
                df_mep_show[["sequencia", "empresa", "lote", "data", "conta", "cc", "tipo", "valor", "historico"]].style.format({
                    "valor": "R$ {:,.2f}"
                }),
                use_container_width=True,
                hide_index=True
            )

            # Export to MXM template
            col_b1, col_b2 = st.columns([1.5, 3])
            with col_b1:
                with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_mep:
                    tmp_mep_path = tmp_mep.name

                try:
                    write_to_excel_template(TEMPLATE_FILE, tmp_mep_path, mep_entries)
                    with open(tmp_mep_path, "rb") as f_mep:
                        mep_bytes = f_mep.read()
                    os.unlink(tmp_mep_path)

                    st.download_button(
                        label="📥 Baixar Lote de MEP no Layout MXM (.xlsx)",
                        data=mep_bytes,
                        file_name=f"Lancamentos_MEP_Empresa_0002_{mes_mep.replace('/', '_')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                except Exception as e:
                    st.error(f"Erro ao gerar planilha MXM: {e}")

    # -------------------------------------------------------------
    # TAB 3: Checklist Mensal de Fechamento Contábil
    # -------------------------------------------------------------
    with tab_checklist:
        st.markdown("### 📋 Checklist Mensal de Fechamento Contábil")
        st.markdown("Acompanhamento colaborativo das etapas obrigatórias de encerramento mensal para as Empresas 0001, 0002 e 0003.")

        if "fechamento_checklist" not in st.session_state:
            st.session_state.fechamento_checklist = {
                "folha_0001": True,
                "folha_0002": True,
                "folha_0003": True,
                "encargos_inss": True,
                "encargos_fgts": True,
                "imobilizado_deprec": False,
                "bancos_conciliados": False,
                "fornecedores_conciliados": False,
                "intercompany_mutuos": False,
                "mep_holding": False,
                "contas_viradas": False,
                "balancete_validado": False
            }

        tasks = [
            ("folha_0001", "1. Integração de Folha de Pagamento - Empresa 0001 (Via Parque / 1SVP)", "RH & Folha"),
            ("folha_0002", "2. Integração de Folha / Pró-Labore - Empresa 0002 (Holding / 2DIR)", "RH & Folha"),
            ("folha_0003", "3. Integração de Folha de Pagamento - Empresa 0003 (Fábrica / 3FAB)", "RH & Folha"),
            ("encargos_inss", "4. Apuração e Conferência de INSS Patronal e Terceiros (GPS)", "Encargos"),
            ("encargos_fgts", "5. Conferência e Conciliação das Guias de FGTS", "Encargos"),
            ("imobilizado_deprec", "6. Cálculo e Contabilização das Quotas de Depreciação (Ativo)", "Patrimônio"),
            ("bancos_conciliados", "7. Conciliação Bancária das Contas Movimento e Aplicações", "Financeiro"),
            ("fornecedores_conciliados", "8. Confronto de Fornecedores e Contas a Pagar", "Contábil"),
            ("intercompany_mutuos", "9. Batimento e Conciliação de Saldos Intercompany (0001 x 0002 x 0003)", "Societário"),
            ("mep_holding", "10. Apuração e Lançamento de Equivalência Patrimonial (MEP) na Holding", "Societário"),
            ("contas_viradas", "11. Auditoria e Reclassificação de Contas Viradas no Balancete", "Auditoria"),
            ("balancete_validado", "12. Fechamento Final e Validação do Balancete Contábil Mensal", "Governança")
        ]

        completed_count = sum(1 for k, _, _ in tasks if st.session_state.fechamento_checklist.get(k, False))
        pct_progress = int((completed_count / len(tasks)) * 100)

        st.progress(pct_progress / 100.0)
        st.markdown(f"**Progresso Geral do Fechamento:** `{completed_count} de {len(tasks)} etapas concluídas ({pct_progress}%)`")
        st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)

        for task_key, task_desc, task_tag in tasks:
            col_t1, col_t2 = st.columns([5, 1])
            with col_t1:
                chk = st.checkbox(
                    task_desc,
                    value=st.session_state.fechamento_checklist.get(task_key, False),
                    key=f"chk_task_{task_key}"
                )
                st.session_state.fechamento_checklist[task_key] = chk
            with col_t2:
                st.markdown(f"<span class='tag-chip' style='margin-top: 5px;'>{task_tag}</span>", unsafe_allow_html=True)


def render_module_placeholder(page_key):
    details = {
        "ativo": {
            "title": "Ativo Imobilizado & Depreciação",
            "icon": "🏭",
            "desc": "Gestão patrimonial do parque fabril e investimentos em instalações das lojas da rede Monte Carlo.",
            "items": [
                "Controle de aquisições e baixas de bens do ativo imobilizado",
                "Cálculo automatizado das quotas mensais de depreciação por centro de custo (3FAB, 1SVP, Lojas)",
                "Exportação do lote contábil de depreciação direto para o MXM",
                "Relatório auxiliar de conciliação do razão do Imobilizado"
            ]
        }
    }
    mod = details.get(page_key, {
        "title": "Módulo em Planejamento",
        "icon": "⚙️",
        "desc": "Rotina contábil em fase de concepção.",
        "items": []
    })
    
    # Top bar
    col_top_back, col_top_title = st.columns([1.5, 4])
    with col_top_back:
        if st.button("⬅ Hub de Rotinas", key=f"btn_back_home_{page_key}"):
            st.session_state.current_page = "home"
            st.rerun()
    with col_top_title:
        st.markdown(f"<div style='font-size: 0.88rem; color: #64748b; padding-top: 6px;'><b>Portal Contábil Monte Carlo</b> &nbsp;/&nbsp; <span>{mod['title']}</span></div>", unsafe_allow_html=True)
    
    st.markdown(f"## {mod['icon']} {mod['title']}")
    st.markdown(f"<p style='color: #475569; font-size: 1.05rem;'>{mod['desc']}</p>", unsafe_allow_html=True)
    
    st.markdown("""
    <div style='background: #fffbeb; border: 1px solid #fde68a; border-radius: 10px; padding: 18px; margin-bottom: 25px;'>
        <b style='color: #b45309;'>🟡 Módulo em Planejamento / Desenvolvimento</b><br>
        <span style='color: #78350f; font-size: 0.9rem;'>
            Esta rotina faz parte do roadmap de expansão do Portal Contábil Monte Carlo. 
            Em breve estará disponível com parametrização completa e integração direta ao ERP MXM.
        </span>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("### 📋 Escopo e Funcionalidades Planejadas:")
    for item in mod['items']:
        st.markdown(f"- ⏳ {item}")
        
    st.markdown("<br>", unsafe_allow_html=True)
    col_btn1, col_btn2 = st.columns([1.5, 3])
    with col_btn1:
        if st.button("Acessar Integrador de Folha (Módulo Ativo) ➔", key=f"btn_go_folha_from_{page_key}", type="primary"):
            st.session_state.current_page = "folha"
            st.rerun()
    with col_btn2:
        if st.button("Voltar à Página Inicial (Hub)", key=f"btn_go_home_from_{page_key}"):
            st.session_state.current_page = "home"
            st.rerun()


# -------------------------------------------------------------
# Global Navigation & Main Router
# -------------------------------------------------------------

# Initialize current_page state
if 'current_page' not in st.session_state:
    st.session_state.current_page = 'home'

# Global Sidebar Header & Navigation
with st.sidebar:
    if os.path.exists("Logo.png"):
        st.image("Logo.png", use_container_width=True)
    else:
        st.image("https://img.icons8.com/fluency/96/000000/accounting.png", width=80)
        
    st.markdown("<div style='text-align: center; font-weight: 700; color: #0b0b0b; margin-top: 5px; margin-bottom: 15px; font-size: 0.95rem; letter-spacing: 0.05em;'>PORTAL CONTÁBIL</div>", unsafe_allow_html=True)
    
    nav_options = [
        ("home", "🏠 Início (Hub de Rotinas)"),
        ("folha", "📊 Folha de Pagamento (MXM)"),
        ("fechamento", "📋 Fechamento Contábil"),
        ("ativo", "🏭 Ativo Imobilizado")
    ]
    nav_keys = [k for k, _ in nav_options]
    nav_labels = {k: label for k, label in nav_options}
    
    current_page = st.session_state.get('current_page', 'home')
    if current_page not in nav_keys:
        current_page = 'home'
    default_nav_idx = nav_keys.index(current_page)
    
    st.markdown("<p style='font-size: 0.72rem; font-weight: 700; color: #94a3b8; text-transform: uppercase; letter-spacing: 0.08em; margin-bottom: 8px;'>MÓDULOS DE ROTINAS</p>", unsafe_allow_html=True)
    sel_page = st.radio(
        "Menu de Módulos",
        options=nav_keys,
        format_func=lambda x: nav_labels[x],
        index=default_nav_idx,
        key=f"sidebar_nav_{current_page}",
        label_visibility="collapsed"
    )
    if sel_page != current_page:
        st.session_state.current_page = sel_page
        st.rerun()

    st.markdown("---")
    
    if st.session_state.current_page == "home":
        st.markdown("**Sobre o Portal:**\n"
                    "Ambiente unificado de automações da **Equipe Contábil Monte Carlo**, integrando relatórios do Alterdata diretamente ao layout do ERP MXM.")
        st.markdown("---")
        last_update_text = get_last_github_update()
        st.markdown(f"""
        <div style='text-align: center; color: #64748b; font-size: 0.78rem; margin-top: 15px;'>
            <b style='color: #475569;'>Equipe Contábil Monte Carlo</b><br>
            <span style='color: #64748b;'>Última atualização (GitHub):<br><b style='color: #0b0b0b;'>{last_update_text}</b></span>
        </div>
        """, unsafe_allow_html=True)
    elif st.session_state.current_page != "folha":
        st.markdown("**Status do Módulo:**\n"
                    "Rotina contábil em planejamento e desenvolvimento para a Equipe Monte Carlo.")
        st.markdown("---")
        last_update_text = get_last_github_update()
        st.markdown(f"""
        <div style='text-align: center; color: #64748b; font-size: 0.78rem; margin-top: 15px;'>
            <b style='color: #475569;'>Equipe Contábil Monte Carlo</b><br>
            <span style='color: #64748b;'>Última atualização (GitHub):<br><b style='color: #0b0b0b;'>{last_update_text}</b></span>
        </div>
        """, unsafe_allow_html=True)

# Early exit router for Home and Placeholders
if st.session_state.current_page == "home":
    render_home_page()
    last_update_text = get_last_github_update()
    st.markdown(f"""
    <div class="footer">
        <div style="font-weight: 600; color: #475569; font-size: 0.95rem; margin-bottom: 4px;">
            Equipe Contábil Monte Carlo
        </div>
        <div style="font-size: 0.82rem; color: #64748b;">
            Última atualização no GitHub: <b style="color: #0b0b0b;">{last_update_text}</b>
        </div>
    </div>
    """, unsafe_allow_html=True)
    st.stop()
elif st.session_state.current_page == "fechamento":
    render_fechamento_contabil()
    last_update_text = get_last_github_update()
    st.markdown(f"""
    <div class="footer">
        <div style="font-weight: 600; color: #475569; font-size: 0.95rem; margin-bottom: 4px;">
            Equipe Contábil Monte Carlo
        </div>
        <div style="font-size: 0.82rem; color: #64748b;">
            Última atualização no GitHub: <b style="color: #0b0b0b;">{last_update_text}</b>
        </div>
    </div>
    """, unsafe_allow_html=True)
    st.stop()
elif st.session_state.current_page != "folha":
    render_module_placeholder(st.session_state.current_page)
    last_update_text = get_last_github_update()
    st.markdown(f"""
    <div class="footer">
        <div style="font-weight: 600; color: #475569; font-size: 0.95rem; margin-bottom: 4px;">
            Equipe Contábil Monte Carlo
        </div>
        <div style="font-size: 0.82rem; color: #64748b;">
            Última atualização no GitHub: <b style="color: #0b0b0b;">{last_update_text}</b>
        </div>
    </div>
    """, unsafe_allow_html=True)
    st.stop()

# --- CONTINUES INTO FOLHA DE PAGAMENTO ---


# Breadcrumb & Module Header
col_top_back, col_top_title = st.columns([1.5, 4])
with col_top_back:
    if st.button("⬅ Hub de Rotinas", key="btn_back_home_folha"):
        st.session_state.current_page = "home"
        st.rerun()
with col_top_title:
    st.markdown("<div style='font-size: 0.88rem; color: #64748b; padding-top: 6px;'><b>Portal Contábil Monte Carlo</b> &nbsp;/&nbsp; <span style='color: #0b0b0b; font-weight: 600;'>Integrador MXM - Folha de Pagamento</span></div>", unsafe_allow_html=True)

# Main Layout
st.title("📊 Geração de Lançamentos Contábeis ERP MXM")
st.markdown("Converta PDFs de Resumo Geral Consolidado de Folha em lotes contábeis importáveis em instantes.")

# Step 1: Upload PDF File
# Step 1: Upload PDF File
st.markdown("### 1. Selecionar Arquivo de Folha (PDF)")
col_up_file, col_up_btn = st.columns([4, 1.2])
with col_up_file:
    uploaded_pdf = st.file_uploader("Arraste e solte o PDF do Resumo Consolidado de Folha de Pagamento", type=["pdf"])
with col_up_btn:
    st.markdown("<div style='height: 28px;'></div>", unsafe_allow_html=True)
    if st.button("🔄 Reanalisar PDF", help="Limpa o cache da sessão e força a reanálise completa do PDF"):
        for k in list(st.session_state.keys()):
            if k.startswith("parsed_pdf_") or k.startswith("pdf_date_applied_"):
                del st.session_state[k]
        st.rerun()

pdf_company_code = None
meta = None
events = None
bases = None
pdf_company_code = None

if uploaded_pdf is not None:
    # Cache parsed result in session state using name and size to avoid re-running on every interaction
    pdf_key = f"parsed_pdf_{uploaded_pdf.name}_{uploaded_pdf.size}"
    
    # Invalidate cache if it contains 0 events
    if pdf_key in st.session_state:
        cached_filiais = st.session_state[pdf_key] or {}
        tot_evs = sum(len((f or {}).get('events') or []) for f in cached_filiais.values())
        if tot_evs == 0:
            del st.session_state[pdf_key]
            
    if pdf_key not in st.session_state:
        with st.spinner("Analisando PDF de folha de pagamento (múltiplas filiais)..."):
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_file:
                uploaded_pdf.seek(0)
                tmp_file.write(uploaded_pdf.getvalue())
                tmp_pdf_path = tmp_file.name
                
            try:
                # Parse PDF using the parser module
                parsed_data = parse_payroll_pdf(tmp_pdf_path)
                # Clean up temp file
                os.unlink(tmp_pdf_path)
                st.session_state[pdf_key] = parsed_data.get('filiais', {})
            except Exception as e:
                import traceback
                tb_str = traceback.format_exc()
                st.error(f"Erro ao analisar o PDF: {str(e)}")
                st.code(tb_str)
                if 'parsed_data' in locals():
                    st.write("Chaves retornadas pelo parser:", list(parsed_data.keys()) if isinstance(parsed_data, dict) else type(parsed_data))
                st.stop()
                
    filiais_dict = st.session_state.get(pdf_key, {})
    if not filiais_dict:
        st.warning("⚠️ Nenhuma filial ou evento contábil foi encontrado no PDF. Verifique se o arquivo enviado é o relatório 'Resumo Geral Consolidado' no formato Alterdata.")
        st.stop()
        
    # Extra safety: if filiais were matched but zero events were parsed, purge corrupted cache and re-run
    tot_events_check = sum(len((f or {}).get('events') or []) for f in filiais_dict.values())
    if tot_events_check == 0:
        if pdf_key in st.session_state:
            del st.session_state[pdf_key]
        st.warning("⚠️ O relatório PDF carregado foi reanalisado porque nenhum evento havia sido indexado no cache. Recarregando...")
        st.rerun()
        
    filial_options = list(filiais_dict.keys())
    
    # Get reference metadata from first filial
    ref_filial = filial_options[0]
    meta = (filiais_dict[ref_filial] or {}).get('metadata') or {}
    pdf_company_code = meta.get('empresa_codigo', '')
       
    # Automatically update date to end of period if not set
    date_key = f"pdf_date_applied_{pdf_key}"
    if date_key not in st.session_state and meta.get('periodo_fim'):
        st.session_state.data_lancamento = meta['periodo_fim']
        st.session_state[date_key] = True
        st.rerun()

# Sidebar Navigation / Details
with st.sidebar:
    if os.path.exists("Logo.png"):
        st.image("Logo.png", use_container_width=True)
    else:
        st.image("https://img.icons8.com/fluency/96/000000/accounting.png", width=80)
        st.title("Integrador Contábil")
    st.markdown("---")
    st.markdown("### Configurações de Importação")
    
    # Defaults configuration
    company_options = ["0001", "0002", "0003"]
    default_index = 0
    if pdf_company_code:
        emp_nome_upper = (meta or {}).get("empresa_nome", "").upper()
        cnpj_clean = re.sub(r"\D", "", (meta or {}).get("cnpj", ""))
        if "JASPER" in emp_nome_upper or "PARTICIPA" in emp_nome_upper or "HOLDING" in emp_nome_upper or cnpj_clean.startswith("09436824"):
            candidate_code = "0002"
        elif pdf_company_code in ("03000", "0003") or "INDUSTRIA" in emp_nome_upper:
            candidate_code = "0003"
        elif pdf_company_code in ("03050", "0001") or "VIA PARQUE" in emp_nome_upper:
            candidate_code = "0001"
        elif pdf_company_code in ("03001", "0002") or "MONTE CARLO JOIAS" in emp_nome_upper:
            candidate_code = "0002"
        else:
            last_3 = pdf_company_code[-3:]
            candidate_code = f"0{last_3}"
            
        if candidate_code in company_options:
            default_index = company_options.index(candidate_code)
        
    company_name_map = {
        "0001": "0001 - Monte Carlo Joias (Via Parque)",
        "0002": "0002 - Jasper Forest Participações (Holding)",
        "0003": "0003 - MC Indústria de Joias (Fábrica)"
    }
    company_code_override = st.selectbox(
        "Código Empresa (MXM)", 
        options=company_options,
        index=default_index,
        format_func=lambda x: company_name_map.get(x, x),
        help="Selecione o código da empresa (0001, 0002 ou 0003) para gerar os lançamentos."
    )
    
    detected_proc = (meta or {}).get("tipo_processo", "folha")
    proc_options = ["folha", "rescisao", "ferias"]
    proc_labels = {
        "folha": "📄 Folha de Pagamento",
        "rescisao": "🚪 Rescisão",
        "ferias": "🏖️ Férias"
    }
    default_proc_idx = proc_options.index(detected_proc) if detected_proc in proc_options else 0
    selected_proc = st.selectbox(
        "Tipo de Processo",
        options=proc_options,
        index=default_proc_idx,
        format_func=lambda x: proc_labels.get(x, x),
        help="Tipo de folha a ser processada. O sistema detecta automaticamente se for Rescisão ou Férias."
    )
    
    ref_mes = ""
    if meta and meta.get("periodo_referencia"):
        try:
            ref_mes = meta.get("periodo_referencia").split("/")[0].zfill(2)
        except Exception:
            pass
    prefix_lote = "RESC" if selected_proc == "rescisao" else ("FER" if selected_proc == "ferias" else "FOLH")
    default_lote = f"{prefix_lote}{ref_mes}" if ref_mes else f"{prefix_lote}XX"

    lote_contabil = st.text_input("Lote Contábil", value=default_lote, max_chars=6, help="Número do lote dos lançamentos contábeis (máx. 6 caracteres).")
    documento_id = st.text_input("Identificador Documento", value=default_lote, max_chars=6, help="Texto curto gravado no campo documento (máx. 6 caracteres).")
    
    # Initialize default date in session state if not present
    if 'data_lancamento' not in st.session_state:
        st.session_state.data_lancamento = datetime.now().strftime("%d/%m/%Y")
        
    date_input_str = st.text_input(
        "Data do Lançamento (DD/MM/AAAA)", 
        value=st.session_state.data_lancamento, 
        help="Digite a data do lançamento contábil no formato DD/MM/AAAA."
    )
    # Save the input to session state to prevent Streamlit widget lock-up
    st.session_state.data_lancamento = date_input_str
    
    st.markdown("---")
    st.markdown("**Sobre o Sistema:**\n"
                "Lê o relatório PDF de Folha de Pagamento, cruza com a parametrização de contas (De-Para) e preenche automaticamente o layout de importação do MXM.")
    
    last_update_text = get_last_github_update()
    st.markdown(f"""
    <div style='text-align: center; color: #64748b; font-size: 0.78rem; margin-top: 25px; padding-top: 15px; border-top: 1px solid #e2ded7;'>
        <b style='color: #475569;'>Equipe Contábil Monte Carlo</b><br>
        <span style='color: #64748b;'>Última atualização (GitHub):<br><b style='color: #0b0b0b;'>{last_update_text}</b></span>
    </div>
    """, unsafe_allow_html=True)

# Construct company and process-specific config file path
if selected_proc == "rescisao":
    CONFIG_FILE = os.path.join(BASE_DIR, f"config_mapping_{company_code_override}_rescisao.json")
    if not os.path.exists(CONFIG_FILE) and os.path.exists(os.path.join(BASE_DIR, f"config_mapping_{company_code_override}.json")):
        CONFIG_FILE = os.path.join(BASE_DIR, f"config_mapping_{company_code_override}.json")
elif selected_proc == "ferias":
    CONFIG_FILE = os.path.join(BASE_DIR, f"config_mapping_{company_code_override}_ferias.json")
    if not os.path.exists(CONFIG_FILE) and os.path.exists(os.path.join(BASE_DIR, f"config_mapping_{company_code_override}.json")):
        CONFIG_FILE = os.path.join(BASE_DIR, f"config_mapping_{company_code_override}.json")
else:
    CONFIG_FILE = os.path.join(BASE_DIR, f"config_mapping_{company_code_override}.json")

# Helper to find column names flexibly
def find_matching_column(columns, keywords):
    for col in columns:
        c_clean = str(col).lower().replace('á', 'a').replace('é', 'e').replace('í', 'i').replace('ó', 'o').replace('ú', 'u').replace('ç', 'c').strip()
        for kw in keywords:
            if kw in c_clean:
                return col
    return None

# Helper to load mapping with fallback to default template if company file does not exist
def load_mapping_from_excel(excel_path):
    import pandas as pd
    excel_mapping = {}
    if os.path.exists(excel_path):
        try:
            df = pd.read_excel(excel_path)
            code_col = find_matching_column(df.columns, ['codigo', 'cod', 'evento'])
            debit_col = find_matching_column(df.columns, ['debito', 'deb'])
            credit_col = find_matching_column(df.columns, ['credito', 'cred'])
            
            if code_col and debit_col and credit_col:
                for _, row in df.iterrows():
                    raw_code = str(row[code_col]).strip()
                    if raw_code.replace(".0", "").isdigit():
                        code = raw_code.replace(".0", "").zfill(3)
                    else:
                        code = raw_code
                        
                    debit = str(row[debit_col]).strip() if pd.notna(row[debit_col]) else ""
                    credit = str(row[credit_col]).strip() if pd.notna(row[credit_col]) else ""
                    
                    if debit.endswith(".0"): debit = debit[:-2]
                    if credit.endswith(".0"): credit = credit[:-2]
                    if debit.lower() in ("nan", "none"): debit = ""
                    if credit.lower() in ("nan", "none"): credit = ""
                    
                    excel_mapping[code] = {
                        "debit_account": debit,
                        "credit_account": credit,
                        "cost_center": ""
                    }
        except Exception as e:
            print("Error loading excel mapping:", e)
    return excel_mapping

def sync_save_mapping(new_mapping, json_path, company_code):
    """Saves mapping to JSON and synchronizes local Excel file on disk."""
    import time
    success = save_mapping(new_mapping, json_path)
    if "rescisao" in json_path:
        excel_path = os.path.join(BASE_DIR, f"De-Para_Empresa_{company_code}_Rescisao.xlsx")
    elif "ferias" in json_path:
        excel_path = os.path.join(BASE_DIR, f"De-Para_Empresa_{company_code}_Ferias.xlsx")
    else:
        excel_path = os.path.join(BASE_DIR, f"De-Para_Empresa_{company_code}.xlsx")
    try:
        export_rows = []
        for code, ev_map in new_mapping.items():
            export_rows.append({
                "Código": code,
                "Conta Débito": ev_map.get("debit_account", ""),
                "Conta Crédito": ev_map.get("credit_account", "")
            })
        pd.DataFrame(export_rows).to_excel(excel_path, index=False, sheet_name="De-Para")
        # Keep json mtime slightly ahead so it won't trigger re-import immediately
        now = time.time()
        os.utime(json_path, (now + 2, now + 2))
    except Exception as e:
        print("Error syncing excel mapping on disk:", e)
    return success

def get_company_mapping(company_code, specific_file):
    mapping = {}
    if "rescisao" in specific_file:
        excel_path = os.path.join(BASE_DIR, f"De-Para_Empresa_{company_code}_Rescisao.xlsx")
    elif "ferias" in specific_file:
        excel_path = os.path.join(BASE_DIR, f"De-Para_Empresa_{company_code}_Ferias.xlsx")
    else:
        excel_path = os.path.join(BASE_DIR, f"De-Para_Empresa_{company_code}.xlsx")
    
    excel_exists = os.path.exists(excel_path)
    json_exists = os.path.exists(specific_file)
    
    # 1. If Excel exists and is newer than JSON (e.g. user edited in Excel on disk), load from Excel
    if excel_exists and json_exists:
        try:
            mtime_excel = os.path.getmtime(excel_path)
            mtime_json = os.path.getmtime(specific_file)
            if mtime_excel > mtime_json + 1:
                mapping = load_mapping_from_excel(excel_path)
                if mapping:
                    # Also preserve charges if already in JSON
                    old_json = load_mapping(specific_file)
                    for chk in ['GPS_PATRONAL', 'GPS_RAT', 'GPS_TERCEIROS', 'FGTS']:
                        if chk not in mapping and chk in old_json:
                            mapping[chk] = old_json[chk]
                    save_mapping(mapping, specific_file)
                    return mapping
        except Exception as e:
            print("Error checking mtime:", e)
        mapping = load_mapping(specific_file)
    elif json_exists:
        mapping = load_mapping(specific_file)
        if mapping and not excel_exists:
            sync_save_mapping(mapping, specific_file, company_code)
    elif excel_exists:
        mapping = load_mapping_from_excel(excel_path)
        if mapping:
            save_mapping(mapping, specific_file)
            
    # 3. Fallbacks if still empty
    if not mapping:
        if company_code in ("0002", "0003"):
            file_0001 = os.path.join(BASE_DIR, "config_mapping_0001.json")
            if os.path.exists(file_0001):
                mapping = load_mapping(file_0001)
                save_mapping(mapping, specific_file)
                sync_save_mapping(mapping, specific_file, company_code)
        else:
            default_file = os.path.join(BASE_DIR, "config_mapping.json")
            if os.path.exists(default_file):
                mapping = load_mapping(default_file)
                
    return mapping
            
if uploaded_pdf is not None and meta is not None:
    # Override company code if configured in sidebar
    if company_code_override:
        meta['empresa_codigo'] = company_code_override
        
    # Display Metadata extracted (Consolidated overview)
    st.markdown(f"""
    <div class="custom-card">
        <h4>📄 Dados Consolidados do Lote de Folha</h4>
        <table style='width: 100%; border-collapse: collapse;'>
            <tr style='border-bottom: 1px solid #f1f5f9;'>
                <td style='padding: 8px; font-weight: 600; color: #64748b; width: 25%;'>Empresa Contábil (MXM):</td>
                <td style='padding: 8px; color: #1e293b;'>{meta.get('empresa_codigo', '')} - {meta.get('empresa_nome', '')}</td>
                <td style='padding: 8px; font-weight: 600; color: #64748b; width: 15%;'>Período:</td>
                <td style='padding: 8px; color: #1e293b;'>{meta.get('periodo_inicio', '')} a {meta.get('periodo_fim', '')}</td>
            </tr>
            <tr>
                <td style='padding: 8px; font-weight: 600; color: #64748b;'>Total de Filiais no PDF:</td>
                <td style='padding: 8px; color: #1e293b; font-weight: 600;'>{len(filiais_dict)} filiais</td>
                <td style='padding: 8px; font-weight: 600; color: #64748b;'>Referência:</td>
                <td style='padding: 8px; color: #1e293b;'>{meta.get('periodo_referencia', '')}</td>
            </tr>
        </table>
    </div>
    """, unsafe_allow_html=True)
    
    # Calcular mudanças de bases e encargos sociais
    rate_patronal = 20.0
    rate_rat = 2.0
    rate_terceiros = 5.8
    rate_fgts = 8.0
    
    # 1. Build dynamic summary by filial
    summary_rows = []
    total_gps_patronal = 0.0
    total_gps_rat = 0.0
    total_gps_terceiros = 0.0
    total_fgts = 0.0
    
    for f_code, filial_data in filiais_dict.items():
        f_meta = (filial_data or {}).get('metadata') or {}
        f_events = (filial_data or {}).get('events') or []
        f_bases = (filial_data or {}).get('bases') or {}
        
        f_base_inss = f_bases.get('inss', 0.0)
        f_base_fgts = f_bases.get('fgts', 0.0)
        f_sc = (filial_data or {}).get('social_charges') or {}
        
        # Calculate dynamic social charges for this filial (prefer pre-calculated from PDF if present)
        sc_pat = f_sc.get('gps_empresa_func', 0.0) + f_sc.get('gps_empresa_socios', 0.0) + f_sc.get('gps_empresa_auton', 0.0)
        sc_rat = f_sc.get('gps_rat', 0.0)
        sc_terc = f_sc.get('gps_terceiros', 0.0)
        sc_fgts = f_sc.get('fgts_total', 0.0)
        
        f_gps_pat = sc_pat if sc_pat > 0 else f_base_inss * (rate_patronal / 100.0)
        f_gps_rt = sc_rat if sc_rat > 0 else f_base_inss * (rate_rat / 100.0)
        f_gps_terc = sc_terc if sc_terc > 0 else f_base_inss * (rate_terceiros / 100.0)
        f_fg_tot = sc_fgts if sc_fgts > 0 else f_base_fgts * (rate_fgts / 100.0)
        
        total_gps_patronal += f_gps_pat
        total_gps_rat += f_gps_rt
        total_gps_terceiros += f_gps_terc
        total_fgts += f_fg_tot
        
        # Proventos vs Descontos
        prov_sum = sum(e['total'] for e in f_events if e.get('type') == 'provento')
        desc_sum = sum(e['total'] for e in f_events if e.get('type') == 'desconto')
        
        # Total accounting flow per filial (Proventos + GPS charges + FGTS)
        f_deb_total = prov_sum + (f_gps_pat + f_gps_rt + f_gps_terc) + f_fg_tot
        
        summary_rows.append({
            "Filial": f_code,
            "Nome da Filial": f_meta.get('empresa_nome', ''),
            "CNPJ": f_meta.get('cnpj', ''),
            "Base INSS": f_base_inss,
            "Base FGTS": f_base_fgts,
            "Total Débitos": f_deb_total,
            "Total Créditos": f_deb_total
        })
        
    df_summary_filiais = pd.DataFrame(summary_rows)
    
    st.markdown("#### 🏢 Resumo de Lançamentos e bases")
    st.dataframe(
        df_summary_filiais.style.format({
            "Base INSS": "R$ {:.2f}",
            "Base FGTS": "R$ {:.2f}",
            "Total Débitos": "R$ {:.2f}",
            "Total Créditos": "R$ {:.2f}"
        }),
        use_container_width=True,
        hide_index=True
    )
    
    # Load existing mapping from config file with fallback
    mapping = get_company_mapping(company_code_override, CONFIG_FILE)
    
    # Step 2: Configure De-Para accounting mapping
    st.markdown("### 2. Parametrização Contábil (De-Para)")
    st.markdown("Associe cada evento extraído às contas de Débito e Crédito no seu plano de contas.")
    
    # Build union of unique events across all filiais
    unique_events = {}
    for f_code, filial_data in filiais_dict.items():
        for ev in ((filial_data or {}).get('events') or []):
            code = ev['code']
            val = ev['total']
            if code not in unique_events:
                unique_events[code] = {
                    'code': code,
                    'description': ev['description'],
                    'type': ev['type'],
                    'section': ev['section'],
                    'total': 0.0
                }
            unique_events[code]['total'] += val
            
    # Sort events by code
    sorted_events = sorted(unique_events.values(), key=lambda x: x['code'])
    
    df_events = []
    for ev in sorted_events:
        code = ev['code']
        desc = ev['description']
        total_val = ev['total']
        section = "Sócio/Autônomo" if ev['section'] == 'socios_autonomos' else ev['type'].capitalize()
        
        # Get existing configuration
        ev_map = mapping.get(code, {})
        debit_acc = ev_map.get('debit_account', '')
        credit_acc = ev_map.get('credit_account', '')
        
        df_events.append({
            "Código": code,
            "Descrição": desc,
            "Tipo": section,
            "Valor Total": total_val,
            "Conta Débito": debit_acc,
            "Conta Crédito": credit_acc
        })
        
    # Append special GPS and FGTS charges to the De-Para table (consolidated totals)
    charges_list = [
        ("GPS_PATRONAL", "INSS Patronal (Empresa)", total_gps_patronal),
        ("GPS_RAT", "INSS RAT/FAP", total_gps_rat),
        ("GPS_TERCEIROS", "INSS Terceiros", total_gps_terceiros),
        ("FGTS", "FGTS (Total Apurado)", total_fgts)
    ]
    for code, desc, val in charges_list:
        ev_map = mapping.get(code, {})
        debit_acc = ev_map.get('debit_account', '')
        credit_acc = ev_map.get('credit_account', '')
        
        df_events.append({
            "Código": code,
            "Descrição": desc,
            "Tipo": "Encargos",
            "Valor Total": val,
            "Conta Débito": debit_acc,
            "Conta Crédito": credit_acc
        })
        
    df_editor = pd.DataFrame(df_events)
    
    # Versioning key for data editor to ensure fresh data display on import/reload
    version_key = f"depara_version_{company_code_override}"
    if version_key not in st.session_state:
        st.session_state[version_key] = 0

    # Quick action buttons for De-Para synchronization
    col_act1, col_act2 = st.columns([1, 1])
    with col_act1:
        if st.button("🔄 Recarregar da Planilha Excel Local", key=f"btn_reload_local_{company_code_override}", help=f"Recarrega os dados diretamente do arquivo De-Para_Empresa_{company_code_override}.xlsx no disco"):
            excel_path_local = os.path.join(BASE_DIR, f"De-Para_Empresa_{company_code_override}.xlsx")
            if os.path.exists(excel_path_local):
                reloaded = load_mapping_from_excel(excel_path_local)
                if reloaded:
                    # Preserve charges if present in existing mapping
                    for chk in ['GPS_PATRONAL', 'GPS_RAT', 'GPS_TERCEIROS', 'FGTS']:
                        if chk not in reloaded and chk in mapping:
                            reloaded[chk] = mapping[chk]
                    sync_save_mapping(reloaded, CONFIG_FILE, company_code_override)
                    st.session_state[version_key] += 1
                    st.success("Parametrização recarregada da planilha local com sucesso!")
                    st.rerun()
                else:
                    st.error("Não foi possível carregar mapeamento da planilha local.")
            else:
                st.error(f"Planilha local {excel_path_local} não encontrada.")
    with col_act2:
        if company_code_override != "0001":
            if st.button("📋 Copiar De-Para da Empresa 0001", key=f"btn_copy_0001_{company_code_override}", help="Preenche as contas desta empresa copiando as contas da Empresa 0001"):
                file_0001 = os.path.join(BASE_DIR, "config_mapping_0001.json")
                if os.path.exists(file_0001):
                    map_0001 = load_mapping(file_0001)
                    if map_0001:
                        sync_save_mapping(map_0001, CONFIG_FILE, company_code_override)
                        st.session_state[version_key] += 1
                        st.success("Mapeamento copiado da Empresa 0001 com sucesso!")
                        st.rerun()

    # Streamlit Data Editor (editable grid)
    edited_df = st.data_editor(
        df_editor,
        column_config={
            "Código": st.column_config.TextColumn("Código", disabled=True, width="medium"),
            "Descrição": st.column_config.TextColumn("Descrição", disabled=True, width="large"),
            "Tipo": st.column_config.TextColumn("Tipo", disabled=True, width="medium"),
            "Valor Total": st.column_config.NumberColumn("Valor Total", disabled=True, format="R$ %.2f", width="medium"),
            "Conta Débito": st.column_config.TextColumn("Conta Débito (MXM)", width="large"),
            "Conta Crédito": st.column_config.TextColumn("Conta Crédito (MXM)", width="large")
        },
        hide_index=True,
        key=f"data_editor_depara_{company_code_override}_{st.session_state[version_key]}",
        use_container_width=True
    )
    
    # Help text about auto-save
    st.info("💡 As parametrizações contábeis são salvas automaticamente ao pressionar Enter ou clicar fora de uma célula editada.")
    
    # Expander to Import/Export De-Para via Spreadsheet
    with st.expander("📥 Importar / Exportar De-Para via Planilha"):
        col_exp, col_imp = st.columns(2)
        with col_exp:
            st.markdown("**Exportar Parametrização:**\n"
                        "Baixe a planilha do mapeamento atual, faça as edições e suba novamente.")
            
            # Prepare export Excel
            export_data = []
            for code, ev_map in mapping.items():
                # Find desc
                desc = next((ev['description'] for ev in sorted_events if ev['code'] == code), "")
                if not desc:
                    desc = {
                        "GPS_PATRONAL": "INSS Patronal (Empresa)",
                        "GPS_RAT": "INSS RAT/FAP",
                        "GPS_TERCEIROS": "INSS Terceiros",
                        "FGTS": "FGTS (Total Apurado)"
                    }.get(code, "")
                export_data.append({
                    "Código": code,
                    "Descrição": desc,
                    "Conta Débito": ev_map.get('debit_account', ''),
                    "Conta Crédito": ev_map.get('credit_account', '')
                })
            df_export = pd.DataFrame(export_data)
            
            import io
            towrite = io.BytesIO()
            df_export.to_excel(towrite, index=False, sheet_name="De-Para")
            towrite.seek(0)
            xlsx_mapping_bytes = towrite.read()
            
            st.download_button(
                label="📥 Baixar Planilha De-Para Atual",
                data=xlsx_mapping_bytes,
                file_name=f"De-Para_Empresa_{company_code_override}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        with col_imp:
            st.markdown("**Importar Parametrização:**\n"
                        "Suba a planilha editada (com as colunas Código, Conta Débito e Conta Crédito).")
            uploaded_mapping = st.file_uploader("Upload Planilha De-Para", type=["xlsx"], key="upload_mapping_excel")
            if uploaded_mapping is not None:
                st.info("Planilha carregada com sucesso. Clique no botão abaixo para processar e atualizar as contas.")
                if st.button("📥 Processar e Atualizar De-Para"):
                    try:
                        df_uploaded = pd.read_excel(uploaded_mapping)
                        code_c = find_matching_column(df_uploaded.columns, ['codigo', 'cod', 'evento'])
                        deb_c = find_matching_column(df_uploaded.columns, ['debito', 'deb'])
                        cred_c = find_matching_column(df_uploaded.columns, ['credito', 'cred'])
                        if not code_c or not deb_c or not cred_c:
                            st.error("A planilha deve conter colunas identificando: 'Código', 'Conta Débito' e 'Conta Crédito'.")
                        else:
                            new_mapping = dict(mapping)
                            for _, row in df_uploaded.iterrows():
                                raw_code = str(row[code_c]).strip()
                                if raw_code.replace(".0", "").isdigit():
                                    code = raw_code.replace(".0", "").zfill(3)
                                else:
                                    code = raw_code
                                
                                debit = str(row[deb_c]).strip() if pd.notna(row[deb_c]) else ""
                                credit = str(row[cred_c]).strip() if pd.notna(row[cred_c]) else ""
                                
                                if debit.endswith(".0"): debit = debit[:-2]
                                if credit.endswith(".0"): credit = credit[:-2]
                                if debit.lower() in ("nan", "none"): debit = ""
                                if credit.lower() in ("nan", "none"): credit = ""
                                
                                new_mapping[code] = {
                                    "debit_account": debit,
                                    "credit_account": credit,
                                    "cost_center": mapping.get(code, {}).get("cost_center", "")
                                }
                            if sync_save_mapping(new_mapping, CONFIG_FILE, company_code_override):
                                st.session_state[version_key] += 1
                                st.success("Parametrização contábil importada com sucesso!")
                                st.toast("De-Para importado com sucesso! 📥")
                                st.rerun()
                            else:
                                st.error("Erro ao salvar parametrização importada.")
                    except Exception as e:
                        st.error(f"Erro ao processar planilha: {str(e)}")
                    
    # Generate the current mapping from the edited dataframe
    current_mapping = mapping.copy()
    for _, row in edited_df.iterrows():
        code = str(row["Código"]).strip()
        current_mapping[code] = {
            "debit_account": str(row["Conta Débito"]).strip() if pd.notna(row["Conta Débito"]) else "",
            "credit_account": str(row["Conta Crédito"]).strip() if pd.notna(row["Conta Crédito"]) else "",
            "cost_center": mapping.get(code, {}).get("cost_center", "")
        }
        
    # Auto-save changes made in data editor without forced rerun (prevents state race conditions)
    if current_mapping != mapping:
        sync_save_mapping(current_mapping, CONFIG_FILE, company_code_override)
        mapping.update(current_mapping)
        st.toast("Mapeamento salvo automaticamente! 💾")
        
    # Step 3: Run Generation and Preview
    st.markdown("### 3. Validação e Fechamento do Lote")
    
    # Validate and parse date from the sidebar input
    try:
        parsed_date = datetime.strptime(st.session_state.data_lancamento.strip(), "%d/%m/%Y")
        formatted_date = parsed_date.strftime("%d%m%Y")
    except ValueError:
        st.sidebar.error("Formato de data inválido. Por favor digite no formato DD/MM/AAAA (exemplo: 31/05/2026).")
        st.stop()
        
    # Seleção do modo de agrupamento
    modo_agrupamento = st.radio(
        "Modo de Geração dos Lançamentos:",
        options=[
            "Consolidado (1 lançamento por evento com o total da empresa)",
            "Detalhado por Filial (lançamentos individuais de cada filial)"
        ],
        index=0,
        horizontal=True,
        help="Consolidado gera o lote totalizado por evento contábil. Detalhado gera os lançamentos separados por filial com o CC padrão da empresa."
    )
    
    if modo_agrupamento.startswith("Consolidado"):
        # Lançamentos consolidados: um lançamento por evento no total da empresa
        consolidated_parsed_data = {
            'metadata': {
                **meta,
                'empresa_codigo': company_code_override,
                'empresa_codigo_original': company_code_override
            },
            'events': [
                {
                    'code': ev['code'],
                    'description': ev['description'],
                    'total': ev['total'],
                    'section': ev['section'],
                    'type': ev['type']
                }
                for ev in sorted_events
            ],
            'social_charges': {
                'gps_empresa_func': total_gps_patronal,
                'gps_empresa_socios': 0.0,
                'gps_empresa_auton': 0.0,
                'gps_rat': total_gps_rat,
                'gps_terceiros': total_gps_terceiros,
                'fgts_total': total_fgts
            }
        }
        
        entry_results = generate_entries(
            consolidated_parsed_data, 
            current_mapping, 
            batch_number=str(lote_contabil or "1"), 
            entry_date=str(formatted_date or ""), 
            doc_number=str(documento_id or "FOLHA"),
            process_type=str(selected_proc or "folha")
        )
        for ue in entry_results.get('unmapped_events', []):
            if 'filial' not in ue:
                ue['filial'] = 'Consolidado'
    else:
        # Lançamentos detalhados por filial
        consolidated_rows = []
        total_debit = 0.0
        total_credit = 0.0
        unmapped_events = []
        global_sequence = 1
        
        for f_code, filial_data in filiais_dict.items():
            f_meta = (filial_data or {}).get('metadata') or {}
            f_events = (filial_data or {}).get('events') or []
            f_bases = (filial_data or {}).get('bases') or {}
            
            f_base_inss = f_bases.get('inss', 0.0)
            f_base_fgts = f_bases.get('fgts', 0.0)
            f_sc = (filial_data or {}).get('social_charges') or {}
            
            sc_pat = f_sc.get('gps_empresa_func', 0.0) + f_sc.get('gps_empresa_socios', 0.0) + f_sc.get('gps_empresa_auton', 0.0)
            sc_rat = f_sc.get('gps_rat', 0.0)
            sc_terc = f_sc.get('gps_terceiros', 0.0)
            sc_fgts = f_sc.get('fgts_total', 0.0)
            
            f_gps_patronal = sc_pat if sc_pat > 0 else f_base_inss * (rate_patronal / 100.0)
            f_gps_rat = sc_rat if sc_rat > 0 else f_base_inss * (rate_rat / 100.0)
            f_gps_terceiros = sc_terc if sc_terc > 0 else f_base_inss * (rate_terceiros / 100.0)
            f_fgts_total = sc_fgts if sc_fgts > 0 else f_base_fgts * (rate_fgts / 100.0)
            
            f_active_parsed_data = {
                'metadata': {
                    **f_meta,
                    'empresa_codigo': company_code_override,
                    'empresa_codigo_original': f_code
                },
                'events': f_events,
                'social_charges': {
                    'gps_empresa_func': f_gps_patronal,
                    'gps_empresa_socios': 0.0,
                    'gps_empresa_auton': 0.0,
                    'gps_rat': f_gps_rat,
                    'gps_terceiros': f_gps_terceiros,
                    'fgts_total': f_fgts_total
                }
            }
            
            res = generate_entries(
                f_active_parsed_data, 
                current_mapping, 
                batch_number=str(lote_contabil or "1"), 
                entry_date=str(formatted_date or ""), 
                doc_number=str(documento_id or "FOLHA"),
                process_type=str(selected_proc or "folha")
            )
            
            for row in res['rows']:
                row['sequencia'] = global_sequence
                global_sequence += 1
                consolidated_rows.append(row)
                
            total_debit += res['total_debit']
            total_credit += res['total_credit']
            
            for ue in res['unmapped_events']:
                unmapped_events.append({
                    'filial': f_code,
                    'code': ue['code'],
                    'description': ue['description'],
                    'amount': ue['amount']
                })
                
        is_balanced = abs(total_debit - total_credit) < 0.01
        entry_results = {
            'rows': consolidated_rows,
            'total_debit': total_debit,
            'total_credit': total_credit,
            'difference': abs(total_debit - total_credit),
            'is_balanced': is_balanced,
            'unmapped_events': unmapped_events
        }
    
    # Metric Columns
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric("Total de Débitos", f"R$ {entry_results['total_debit']:,.2f}")
    with col2:
        st.metric("Total de Créditos", f"R$ {entry_results['total_credit']:,.2f}")
    with col3:
        st.metric("Diferença", f"R$ {entry_results['difference']:,.2f}")
    with col4:
        st.markdown("Status do Lote:")
        if entry_results['is_balanced']:
            st.markdown('<span class="status-badge status-balanced">Lote Balanceado ✓</span>', unsafe_allow_html=True)
        else:
            st.markdown('<span class="status-badge status-unbalanced">Lote Desbalanceado ✗</span>', unsafe_allow_html=True)
            
    # Calculation of Credit - Debit for accounts starting with 2
    accounts_2 = {}
    for row in entry_results['rows']:
        conta = row['conta']
        # Clean dot for prefix check
        clean_conta = conta.replace(".", "").strip()
        if clean_conta.startswith('2'):
            if conta not in accounts_2:
                accounts_2[conta] = {'debitos': 0.0, 'creditos': 0.0}
            
            if row['tipo'] == 'D':
                accounts_2[conta]['debitos'] += row['valor']
            elif row['tipo'] == 'C':
                accounts_2[conta]['creditos'] += row['valor']
                
    if accounts_2:
        st.markdown("#### ⚖️ Saldos das Contas de Passivo (Iniciadas em 2)")
        st.markdown("Saldos líquidos a recolher / pagar (Créditos - Débitos) gerados no lote para cada conta contábil:")
        
        summary_2 = []
        for conta, vals in accounts_2.items():
            deb = vals['debitos']
            cred = vals['creditos']
            saldo = cred - deb
            summary_2.append({
                "Conta Contábil": conta,
                "Total Débitos (D)": deb,
                "Total Créditos (C)": cred,
                "Saldo Líquido (C - D)": saldo
            })
        
        df_summary_2 = pd.DataFrame(summary_2)
        st.dataframe(
            df_summary_2.style.format({
                "Total Débitos (D)": "R$ {:.2f}",
                "Total Créditos (C)": "R$ {:.2f}",
                "Saldo Líquido (C - D)": "R$ {:.2f}"
            }),
            use_container_width=True,
            hide_index=True
        )
            
    # Unmapped alerts
    if entry_results['unmapped_events']:
        st.warning(f"Atenção: Existem {len(entry_results['unmapped_events'])} eventos com parametrização incompleta (sem conta de Débito e/ou Crédito). Estes não serão exportados.")
        with st.expander("Ver eventos não parametrizados"):
            unmapped_df = pd.DataFrame(entry_results['unmapped_events'])
            st.dataframe(unmapped_df[["filial", "code", "description", "amount"]].rename(columns={
                "filial": "Filial", "code": "Código", "description": "Descrição", "amount": "Valor Total"
            }))
            
    # Export area
    if len(entry_results['rows']) > 0:
        st.markdown("---")
        st.markdown("#### Visualização dos Lançamentos Gerados")
        preview_df = pd.DataFrame(entry_results['rows'])
        st.dataframe(
            preview_df.rename(columns={
                "empresa": "Empresa", "lote": "Lote", "data": "Data", "documento": "Documento",
                "conta": "Conta Contábil", "cc": "Centro Custo", "tipo": "D/C", "historico": "Histórico", "valor": "Valor", "sequencia": "Seq"
            }), 
            use_container_width=True
        )
        
        # Export Action
        st.markdown("### 4. Exportar Layout MXM")
        if not os.path.exists(TEMPLATE_FILE):
            st.error(f"Arquivo modelo do layout MXM não encontrado no caminho esperado:\n`{TEMPLATE_FILE}`\nPor favor, garanta que o arquivo modelo de referência esteja nesta pasta.")
        else:
            # Prepare temporary file for output
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_out:
                output_xlsx_path = tmp_out.name
                
            try:
                write_success = write_to_excel_template(
                    template_path=TEMPLATE_FILE,
                    output_path=output_xlsx_path,
                    entries_data=entry_results['rows']
                )
                
                if write_success:
                    # Load generated bytes for Streamlit download button
                    with open(output_xlsx_path, "rb") as f:
                        xlsx_bytes = f.read()
                        
                    # Clean up temp file
                    os.unlink(output_xlsx_path)
                    
                    file_name = f"Importacao_MXM_Folha_{meta['periodo_referencia'].replace('/', '_')}.xlsx"
                    
                    st.download_button(
                        label="📥 Baixar Planilha Pronta para Importação no MXM",
                        data=xlsx_bytes,
                        file_name=file_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    st.success("Planilha preenchida com sucesso e pronta para download! Basta importá-la no MXM ERP usando a funcionalidade de Importação de Lançamentos em Excel.")
                    
            except Exception as e:
                st.error(f"Erro ao preencher o layout Excel: {str(e)}")
                if os.path.exists(output_xlsx_path):
                    os.unlink(output_xlsx_path)
    else:
        st.error(f"⚠️ Não foram gerados lançamentos contábeis para a Empresa {company_code_override}.")
        st.info("Motivo: Os eventos identificados na folha de pagamento ainda não possuem contas contábeis de Débito e/ou Crédito parametrizadas no De-Para desta empresa.")
        
        events_in_pdf = [ev['code'] for ev in sorted_events]
        mapped_count = sum(1 for c in events_in_pdf if (current_mapping.get(c, {}).get('debit_account') or current_mapping.get(c, {}).get('credit_account')))
        st.markdown(f"**Diagnóstico:** Dos **{len(events_in_pdf)}** eventos encontrados neste PDF, apenas **{mapped_count}** possuem contas configuradas.")
        
        col_diag1, col_diag2 = st.columns(2)
        with col_diag1:
            if company_code_override != "0001":
                if st.button("📋 Copiar Parametrização da Empresa 0001 para esta Empresa", key="btn_copy_from_0001_diag"):
                    file_0001 = os.path.join(BASE_DIR, "config_mapping_0001.json")
                    if os.path.exists(file_0001):
                        map_0001 = load_mapping(file_0001)
                        if map_0001:
                            sync_save_mapping(map_0001, CONFIG_FILE, company_code_override)
                            st.session_state[version_key] = st.session_state.get(version_key, 0) + 1
                            st.success("Parametrização copiada da Empresa 0001 com sucesso!")
                            st.rerun()
        with col_diag2:
            excel_path_local = os.path.join(BASE_DIR, f"De-Para_Empresa_{company_code_override}.xlsx")
            if st.button("🔄 Recarregar da Planilha Excel Local", key="btn_reload_excel_diag"):
                if os.path.exists(excel_path_local):
                    loaded = load_mapping_from_excel(excel_path_local)
                    if loaded:
                        sync_save_mapping(loaded, CONFIG_FILE, company_code_override)
                        st.session_state[version_key] = st.session_state.get(version_key, 0) + 1
                        st.success("De-Para recarregado do Excel com sucesso!")
                        st.rerun()
elif uploaded_pdf is not None and f"parsed_pdf_{uploaded_pdf.name}_{uploaded_pdf.size}" not in st.session_state:
    st.info("Verifique se o PDF selecionado é o relatório 'Resumo Geral Consolidado' no formato de folha correto.")
else:
    # App instructions when no file uploaded
    st.markdown("""
    <div class="custom-card" style='text-align: center; padding: 40px;'>
        <img src="https://img.icons8.com/color/96/000000/pdf.png" width="80" style="margin-bottom: 20px;"/>
        <h3>Nenhum PDF Carregado</h3>
        <p style='color: #64748b;'>Faça o upload do relatório PDF de Resumo Geral Consolidado de Folha de Pagamento acima para iniciar.</p>
    </div>
    """, unsafe_allow_html=True)

# Footer
last_update_text = get_last_github_update()
st.markdown(f"""
<div class="footer">
    <div style="font-weight: 600; color: #475569; font-size: 0.95rem; margin-bottom: 4px;">
        Equipe Contábil Monte Carlo
    </div>
    <div style="font-size: 0.82rem; color: #64748b;">
        Última atualização no GitHub: <b style="color: #0b0b0b;">{last_update_text}</b>
    </div>
</div>
""", unsafe_allow_html=True)
