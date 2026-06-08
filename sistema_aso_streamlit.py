import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
from datetime import datetime, timedelta
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from io import BytesIO
import os
from PIL import Image
import base64
import json

# =========================================================
# CONFIGURAÇÃO DE CAMINHOS E PERSISTÊNCIA LOCAL (JSON)
# =========================================================

BASE_DIR = os.path.dirname(__file__) if "__file__" in locals() else "."
ARQUIVO_AGENDADOS = os.path.join(BASE_DIR, "historico_agendados.json")

def localizar_arquivo(caminho_local, nome_arquivo):
    if os.path.exists(caminho_local):
        return caminho_local
    caminho_projeto = os.path.join(BASE_DIR, nome_arquivo)
    return caminho_projeto if os.path.exists(caminho_projeto) else None

# Caminhos dos arquivos de imagem
PATH_ASO_IMG = localizar_arquivo(r"C:\Users\dilceu.gomes\Desktop\sistema_aso\ASO.png", "ASO.png")
PATH_LOGO = localizar_arquivo(r"C:\Users\dilceu.gomes\Desktop\sistema_aso\logo.png", "logo.png")
PATH_LOGO_DOC = localizar_arquivo(r"C:\Users\dilceu.gomes\Desktop\sistema_aso\adivitta.png", "adivitta.png")

# Funções de leitura e escrita do histórico local
def carregar_agendados_locais():
    if os.path.exists(ARQUIVO_AGENDADOS):
        try:
            with open(ARQUIVO_AGENDADOS, "r", encoding="utf-8") as f:
                return set(json.load(f))
        except:
            return set()
    return set()

def salvar_status_local(nome_colaborador, status_marcado):
    agendados = carregar_agendados_locais()
    if status_marcado:
        agendados.add(nome_colaborador.strip())
    else:
        agendados.discard(nome_colaborador.strip())
    
    with open(ARQUIVO_AGENDADOS, "w", encoding="utf-8") as f:
        json.dump(list(agendados), f, ensure_ascii=False, indent=4)

def carregar_imagem_base64(path):
    if path and os.path.exists(path):
        with open(path, "rb") as img_file:
            return base64.b64encode(img_file.read()).decode()
    return None

# =========================================================
# CONFIGURAÇÃO DA PÁGINA
# =========================================================

try:
    page_icon = Image.open(PATH_ASO_IMG) if PATH_ASO_IMG else "🛡️"
except:
    page_icon = "🛡️"

st.set_page_config(
    page_title="Gestão ASO Pro",
    layout="wide",
    page_icon=page_icon
)

# =========================================================
# CSS GLOBAL E TOPO PERSONALIZADO
# =========================================================

aso_base64 = carregar_imagem_base64(PATH_ASO_IMG)

st.markdown(f"""
<style>
[data-testid="stSidebarNav"] {{display: none !important;}}
.stApp {{ background-color: #f1f5f9; }}
section[data-testid="stSidebar"] {{ background: linear-gradient(180deg, #8390a8, #1e293b); }}
section[data-testid="stSidebar"] label, section[data-testid="stSidebar"] p {{ color: white !important; }}
div[data-testid="metric-container"] {{ background: white; border-radius: 18px; padding: 15px; border: 1px solid #e2e8f0; box-shadow: 0 4px 18px rgba(0,0,0,0.06); }}
.stDownloadButton button {{ width: 100%; background: linear-gradient(90deg, #2563eb, #1d4ed8); color: white; border-radius: 12px; font-weight: bold; }}

section[data-testid="stSidebar"] .stButton > button {{
    background-color: #2563eb !important;
    color: white !important;
    border-radius: 12px !important;
    border: 1px solid rgba(255,255,255,0.3) !important;
    font-weight: bold !important;
    height: 45px !important;
    transition: 0.3s;
}}
section[data-testid="stSidebar"] .stButton > button:hover {{
    background-color: #1d4ed8 !important;
    border: 1px solid #f9cc0b !important;
}}
.footer {{ position: fixed; left: 0; bottom: 0; width: 100%; background: rgba(0,0,0,0.8); color: white; text-align: center; padding: 10px; font-size: 13px; z-index: 999; }}
.menu-titulo {{ color: white; font-weight: bold; font-size: 14px; margin-top: 20px; margin-bottom: 10px; border-bottom: 1px solid rgba(255,255,255,0.2); padding-bottom: 5px; }}
.header-wrapper {{ display: flex; align-items: center; gap: 20px; margin-bottom: 5px; }}
.main-title {{ font-size: 40px; font-weight: 800; color: #111827; margin: 0; }}
.sub-title {{ color: #64748b; margin-bottom: 25px; margin-left: 120px; }}
div[data-baseweb="select"] > div {{ background-color: #ffffff !important; border-radius: 12px !important; border: 1px solid #cbd5e1 !important; }}
div[data-baseweb="select"] > div:focus-within {{ border-color: #2563eb !important; box-shadow: 0 0 0 3px rgba(37, 99, 235, 0.15) !important; }}
</style>

<div class="header-wrapper">
    {"<img src='data:image/png;base64," + aso_base64 + "' width='100'>" if aso_base64 else "🛡️"}
    <h1 class="main-title">Gestão ASO Pro</h1>
</div>
<div class="sub-title">Sistema Inteligente de Gestão de ASO</div>
""", unsafe_allow_html=True)

# =========================================================
# LÓGICA DE DADOS
# =========================================================

SHEET_ID = "1G_oVT9gK-n_jGh5R4g65qUwK_MfQGvCX-SA4NHNNflU"
UNIDADES = {
    "D-ITU": "1323067532", "D-MG": "1071212860", 
    "J-CHP": "1549718037", "S-SJ": "1712391604", "J-CTBA": "145843404"
}

SHEET_ID_GLOBAL = "1y98U3eK7JXJqQaMC0i7eFbwpvp97Nuyeml5Dis0UCUg"
GID_COLABORADORES = "595994340"

@st.cache_data(ttl=60) # Ajustado para 1 minuto para atualizar rápido
def load_data(gid):
    url = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid={gid}"
    df = pd.read_csv(url)
    df.columns = df.columns.str.strip()
    return df

def gerar_docx(dados, tipo, data_sugestao):
    doc = Document()
    if PATH_LOGO_DOC and os.path.exists(PATH_LOGO_DOC):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run().add_picture(PATH_LOGO_DOC, width=Inches(1.2))

    titulo = doc.add_paragraph()
    titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = titulo.add_run(f"FORMULÁRIO PARA AGENDAMENTO {tipo}")
    run.bold, run.font.size = True, Pt(16)

    table = doc.add_table(rows=7, cols=2)
    table.style = "Table Grid"
    labels = ["Nome Completo", "Cargo", "Setor", "Unidade", "Cliente", "Local", "Data Sugestão"]
    valores = [str(dados["Nome"]), str(dados["Cargo"]), str(dados["Setor"]), str(dados["UNIDADE"]), "MACROMAQ", "Arapoti", data_sugestao.strftime("%d/%m/%Y")]

    for i, (l, v) in enumerate(zip(labels, valores)):
        table.rows[i].cells[0].text, table.rows[i].cells[1].text = l, v
        table.rows[i].cells[0].paragraphs[0].runs[0].bold = True

    target = BytesIO()
    doc.save(target)
    return target.getvalue()

# =========================================================
# SIDEBAR
# =========================================================

if PATH_LOGO and os.path.exists(PATH_LOGO):
    st.sidebar.image(PATH_LOGO, width=300)

aba_nome = st.sidebar.selectbox(
    "🏢 Selecione a unidade",
    list(UNIDADES.keys())
)

st.sidebar.success("✅ Sistema Online")
st.sidebar.markdown("<br>", unsafe_allow_html=True)
st.sidebar.markdown("""<div class='menu-titulo'>⚡ ACESSO RÁPIDO</div>""", unsafe_allow_html=True)

if st.sidebar.button("📋 CHECKLIST SSMA"):
    st.switch_page("pages/app_ssma_ia.py")

# =========================================================
# PROCESSAMENTO PRINCIPAL
# =========================================================
try:
    # 1. Carrega os agendados salvos do arquivo físico estável (.json)
    lista_agendados_salvos = carregar_agendados_locais()

    # 2. Carrega dados da Planilha 1
    df = load_data(UNIDADES[aba_nome])
    
    # 3. Carrega dados da Planilha 2
    url_global = f"https://docs.google.com/spreadsheets/d/{SHEET_ID_GLOBAL}/export?format=csv&gid={GID_COLABORADORES}"
    df_global_raw = pd.read_csv(url_global, header=None)
    
    df_global = pd.DataFrame()
    df_global["Nome"] = df_global_raw[8].astype(str).str.strip()      
    df_global["Cargo"] = df_global_raw[5].astype(str).str.strip()     
    df_global["Setor"] = df_global_raw[4].astype(str).str.strip()     
    df_global["UNIDADE"] = df_global_raw[2].astype(str).str.strip()   
    df_global["Matricula"] = df_global_raw[6].astype(str).str.strip() 
    df_global["CPF"] = df_global_raw[18].astype(str).str.strip()      

    # Tratamento contra valores nulos e cabeçalhos textuais duplicados
    df_global = df_global.dropna(subset=["Nome"])
    df_global = df_global[(df_global["Nome"].str.lower() != "nome") & 
                          (df_global["Nome"].str.lower() != "nan") & 
                          (df_global["Nome"] != "")]

    # ---------------------------------------------------------
    # SEÇÃO: BUSCA ATIVA GLOBAL
    # ---------------------------------------------------------
    st.markdown("### 🔍 Busca Ativa e Emissão de Guia Avulsa")
    
    # Limpeza e ordenação estrita de strings para evitar erro de tipo float vs str
    lista_nomes = sorted([nome for nome in df_global["Nome"].unique() if isinstance(nome, str) and nome.strip() != ""])
    nome_selecionado = st.selectbox("Digite ou selecione o nome do colaborador:", [""] + lista_nomes, index=0)

    if nome_selecionado != "":
        dados_colaborador = df_global[df_global["Nome"] == nome_selecionado].iloc[0]
        colab_nome = dados_colaborador['Nome']
        
        ja_agendado_busca = colab_nome in lista_agendados_salvos
        status_agendado_busca = '<span style="background-color: #2563eb; color: white; padding: 4px 10px; border-radius: 8px; font-size: 14px; margin-left: 15px;">📌 AGENDADO</span>' if ja_agendado_busca else ""
        
        html_busca = f"""
        <div style="background:#f8fafc; border-left:8px solid #3b82f6; border-radius:18px; padding:22px; margin-top:10px; margin-bottom:15px; box-shadow:0 4px 18px rgba(0,0,0,0.05); font-family:Arial;">
            <div style="font-size:22px; font-weight:700; color:#1e3a8a; margin-bottom:10px;">🔍 Registro Encontrado: {colab_nome} {status_agendado_busca}</div>
            <div style="color:#475569; font-size:15px; line-height:1.8;"><b>👔 Cargo:</b> {dados_colaborador['Cargo']}<br><b>🏭 Setor:</b> {dados_colaborador['Setor']}<br><b>🏢 Unidade Base:</b> {dados_colaborador['UNIDADE']}<br><b>🆔 Matrícula:</b> {dados_colaborador['Matricula']}</div>
        </div>"""
        
        components.html(html_busca, height=210)
        
        c_busca1, c_busca2, c_busca3 = st.columns([2, 2, 1])
        with c_busca1: tipo_busca = st.selectbox("Tipo de Exame (Avulso)", ["MUDANÇA DE RISCO", "RETORNO", "PERIÓDICO", "ADMISSIONAL", "DEMISSIONAL"], key="tipo_busca")
        with c_busca2: dt_s_busca = st.date_input("Data sugerida (Avulso)", value=datetime.now() + timedelta(days=2), key="data_busca")
        with c_busca3:
            st.markdown("<br>", unsafe_allow_html=True)
            # KEY baseada no nome do funcionário evita bugs de cache do Streamlit
            foi_agendado_b = st.checkbox("Marcar como Agendado", value=ja_agendado_busca, key=f"chk_b_{colab_nome}")
            
            if foi_agendado_b != ja_agendado_busca:
                salvar_status_local(colab_nome, foi_agendado_b)
                st.rerun()
        
        btn_doc_busca = gerar_docx(dados_colaborador, tipo_busca, dt_s_busca)
        st.download_button(label=f"📥 Baixar Formulário de {colab_nome.split()[0]}", data=btn_doc_busca, file_name=f"ASO_AVULSO_{colab_nome}.docx", key="btn_busca_download")
        st.markdown("---")

    # ---------------------------------------------------------
    # MONITORAMENTO AUTOMÁTICO DE PRAZOS
    # ---------------------------------------------------------
    st.markdown("### 📊 Monitoramento Automático de Prazos (Próximos Vencimentos)")
    if not df.empty:
        hoje = datetime.now()
        df["Venc"] = pd.to_datetime(df["Venc"], dayfirst=True, errors="coerce")
        df = df.dropna(subset=["Venc"])
        df["Dias"] = (df["Venc"] - hoje).dt.days
        alertas = df[df["Venc"] <= hoje + timedelta(days=10)].copy().sort_values(by="Venc")

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("🏢 Unidade", aba_nome); c2.metric("👥 Colaboradores", len(df))
        c3.metric("⚠️ Pendentes", len(alertas)); c4.metric("🚨 Vencidos", len(alertas[alertas["Dias"] < 0]))

        st.markdown("---")
        cols = st.columns(2)
        for idx, row in alertas.reset_index().iterrows():
            col = cols[idx % 2]
            dias = int(row["Dias"])
            nome_alerta = row["Nome"]
            
            ja_marcado_card = nome_alerta in lista_agendados_salvos

            cor, fundo, status = ("#ef4444", "#fff1f2", "🚨 ASO VENCIDO") if dias < 0 else (("#f59e0b", "#fff7ed", f"⚠️ Vence in {dias} dias") if dias <= 3 else ("#10b981", "#ecfdf5", f"✅ Vence in {dias} dias"))
            badge_agendado = '<span style="background-color: #2563eb; color: white; padding: 3px 8px; border-radius: 6px; font-size: 13px; font-weight: bold; margin-left: 10px;">📌 AGENDADO</span>' if ja_marcado_card else ""

            html_card = f"""
            <div style="background:{fundo}; border-left:8px solid {cor}; border-radius:18px; padding:22px; margin-bottom:15px; box-shadow:0 4px 18px rgba(0,0,0,0.08); font-family:Arial;">
                <div style="font-size:22px; font-weight:700; color:#111827; margin-bottom:10px;">{nome_alerta} {badge_agendado}</div>
                <div style="color:#475569; font-size:15px; line-height:1.8;">👔 <b>Cargo:</b> {row['Cargo']}<br>🏭 <b>Setor:</b> {row['Setor']}<br>📅 <b>Vencimento:</b> {row['Venc'].strftime('%d/%m/%Y')}</div>
                <div style="margin-top:15px; font-size:16px; font-weight:bold; color:{cor};">{status}</div>
            </div>"""

            with col:
                components.html(html_card, height=250)
                with st.expander(f"📄 Gerar Agendamento - {nome_alerta.split()[0]}"):
                    c_card1, c_card2 = st.columns([3, 2])
                    with c_card1:
                        tipo = st.selectbox("Tipo de Exame", ["PERIÓDICO", "MUDANÇA DE RISCO", "RETORNO"], key=f"t_{nome_alerta}")
                        dt_s = st.date_input("Data sugerida", value=hoje + timedelta(days=2), key=f"d_{nome_alerta}")
                    with c_card2:
                        st.markdown("<br>", unsafe_allow_html=True)
                        # KEY baseada no nome do alerta impede a perda de estado ao atualizar a unidade
                        foi_agendado_card = st.checkbox("Marcar como Agendado", value=ja_marcado_card, key=f"chk_c_{nome_alerta}")
                        
                        if foi_agendado_card != ja_marcado_card:
                            salvar_status_local(nome_alerta, foi_agendado_card)
                            st.rerun()
                            
                    btn_doc = gerar_docx(row, tipo, dt_s)
                    st.download_button(label="📥 Baixar Documento", data=btn_doc, file_name=f"ASO_{nome_alerta}.docx", key=f"b_{nome_alerta}")
except Exception as e:
    st.error(f"Erro ao processar dados: {e}")

st.markdown("""<div class="footer">© 2026 Gestão ASO´s | Desenvolvido por: Dilceu Junior</div>""", unsafe_allow_html=True)