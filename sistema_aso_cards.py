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
import unicodedata
from urllib.parse import unquote

# =========================================================
# CONFIGURAÇÃO DE CAMINHOS DINÂMICOS
# =========================================================

BASE_DIR = os.path.dirname(__file__) if "__file__" in locals() else "."

def localizar_arquivo(caminho_local, nome_arquivo):
    if os.path.exists(caminho_local):
        return caminho_local
    caminho_projeto = os.path.join(BASE_DIR, nome_arquivo)
    return caminho_projeto if os.path.exists(caminho_projeto) else None

# Caminhos dos arquivos
PATH_ASO_IMG = localizar_arquivo(r"C:\Users\dilceu.gomes\Desktop\sistema_aso\ASO.png", "ASO.png")
PATH_LOGO = localizar_arquivo(r"C:\Users\dilceu.gomes\Desktop\sistema_aso\logo.png", "logo.png")
PATH_LOGO_DOC = localizar_arquivo(r"C:\Users\dilceu.gomes\Desktop\sistema_aso\adivitta.png", "adivitta.png")

def carregar_imagem_base64(path):
    if path and os.path.exists(path):
        with open(path, "rb") as img_file:
            return base64.b64encode(img_file.read()).decode()
    return None

def normalizar_texto(texto):
    if not isinstance(texto, str):
        return ""
    texto = unicodedata.normalize("NFKD", texto).encode("ASCII", "ignore").decode("ASCII")
    return texto.strip().upper()

# =========================================================
# CONFIGURAÇÃO DA PÁGINA
# =========================================================

try:
    if PATH_ASO_IMG:
        page_icon = Image.open(PATH_ASO_IMG)
    else:
        page_icon = "🛡️"
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
/* ESCONDER NAVEGAÇÃO AUTOMÁTICA */
[data-testid="stSidebarNav"] {{display: none !important;}}

/* FUNDO E ESTILOS GERAIS */
.stApp {{ background-color: #f1f5f9; }}
section[data-testid="stSidebar"] {{ background: linear-gradient(180deg, #8390a8, #1e293b); }}
section[data-testid="stSidebar"] label, section[data-testid="stSidebar"] p {{ color: white !important; }}
div[data-testid="metric-container"] {{ background: white; border-radius: 18px; padding: 15px; border: 1px solid #e2e8f0; box-shadow: 0 4px 18px rgba(0,0,0,0.06); }}

/* BOTÃO DE DOWNLOAD */
.stDownloadButton button {{ width: 100%; background: linear-gradient(90deg, #2563eb, #1d4ed8); color: white; border-radius: 12px; font-weight: bold; }}

/* AJUSTE DO BOTÃO NA SIDEBAR */
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

.menu-titulo {{
    color: white;
    font-weight: bold;
    font-size: 14px;
    margin-top: 20px;
    margin-bottom: 10px;
    border-bottom: 1px solid rgba(255,255,255,0.2);
    padding-bottom: 5px;
}}

.header-wrapper {{
    display: flex;
    align-items: center;
    gap: 20px;
    margin-bottom: 5px;
}}
.main-title {{
    font-size: 40px;
    font-weight: 800;
    color: #111827;
    margin: 0;
}}
.sub-title {{
    color: #64748b;
    margin-bottom: 25px;
    margin-left: 120px;
}}
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

@st.cache_data(ttl=300)
def load_data(gid):
    try:
        url = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid={gid}"
        df = pd.read_csv(url)
        df.columns = df.columns.str.strip()
        return df
    except:
        return pd.DataFrame()

def gerar_docx(dados, tipo, data_sugestao, unidade_padrao="MACROMAQ"):
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
    valores = [
        str(dados.get("Nome", "")),
        str(dados.get("Cargo", "")),
        str(dados.get("Setor", "")),
        str(dados.get("UNIDADE", unidade_padrao)),
        "MACROMAQ",
        "Arapoti",
        data_sugestao.strftime("%d/%m/%Y")
    ]

    for i, (l, v) in enumerate(zip(labels, valores)):
        table.rows[i].cells[0].text, table.rows[i].cells[1].text = l, v
        table.rows[i].cells[0].paragraphs[0].runs[0].bold = True

    target = BytesIO()
    doc.save(target)
    return target.getvalue()

# =========================================================
# CAPTURA INTELIGENTE DE URL (?colaborador=...)
# =========================================================
colab_param = st.query_params.get("colaborador", None)

nome_alvo_encontrado = None
unidade_correta = None

if colab_param:
    nome_url_norm = normalizar_texto(unquote(str(colab_param)))
    # Varre as 5 unidades para descobrir a filial correta do colaborador
    for nome_u, gid_u in UNIDADES.items():
        df_u = load_data(gid_u)
        if not df_u.empty and "Nome" in df_u.columns:
            for nome_original in df_u["Nome"].dropna().astype(str).unique():
                if normalizar_texto(nome_original) == nome_url_norm or nome_url_norm in normalizar_texto(nome_original):
                    unidade_correta = nome_u
                    nome_alvo_encontrado = nome_original
                    break
        if unidade_correta:
            break

# Se achou uma filial diferente via URL, atualiza o seletor da Sidebar automaticamente
if unidade_correta:
    if "sidebar_unidade" not in st.session_state or st.session_state.sidebar_unidade != unidade_correta:
        st.session_state["sidebar_unidade"] = unidade_correta

# =========================================================
# SIDEBAR
# =========================================================

if PATH_LOGO and os.path.exists(PATH_LOGO):
    st.sidebar.image(PATH_LOGO, width=300)

lista_unidades = list(UNIDADES.keys())
idx_unidade_padrao = lista_unidades.index(unidade_correta) if unidade_correta in lista_unidades else 0

aba_nome = st.sidebar.selectbox(
    "🏢 Selecione a unidade",
    lista_unidades,
    index=idx_unidade_padrao,
    key="sidebar_unidade"
)

st.sidebar.success("✅ Sistema Online")

st.sidebar.markdown("<br>", unsafe_allow_html=True)
st.sidebar.markdown(
    "<div class='menu-titulo'>⚡ ACESSO RÁPIDO</div>",
    unsafe_allow_html=True
)

if st.sidebar.button("📋 CHECKLIST SSMA"):
    st.switch_page("pages/app_ssma_ia.py")

# =========================================================
# PROCESSAMENTO PRINCIPAL
# =========================================================
try:
    df = load_data(UNIDADES[aba_nome])
    if not df.empty and "Nome" in df.columns:
        hoje = datetime.now()
        
        # -------------------------------------------------------------
        # 1. BUSCA ATIVA E EMISSÃO DE GUIA AVULSA
        # -------------------------------------------------------------
        st.markdown("### 🔍 Busca Ativa e Emissão de Guia Avulsa")
        
        lista_colaboradores_unid = sorted(df["Nome"].dropna().astype(str).unique().tolist())
        
        # Define o colaborador selecionado
        idx_selecao = 0
        if nome_alvo_encontrado and nome_alvo_encontrado in lista_colaboradores_unid:
            idx_selecao = lista_colaboradores_unid.index(nome_alvo_encontrado)
            opcoes_select = lista_colaboradores_unid
        elif colab_param:
            # Tenta encontrar por aproximação dentro da unidade atual
            idx_encontrado = None
            nome_url_norm = normalizar_texto(unquote(str(colab_param)))
            for i, n in enumerate(lista_colaboradores_unid):
                if normalizar_texto(n) == nome_url_norm or nome_url_norm in normalizar_texto(n):
                    idx_encontrado = i
                    break
            if idx_encontrado is not None:
                idx_selecao = idx_encontrado
                opcoes_select = lista_colaboradores_unid
            else:
                opcoes_select = ["Selecione um colaborador..."] + lista_colaboradores_unid
                idx_selecao = 0
        else:
            opcoes_select = ["Selecione um colaborador..."] + lista_colaboradores_unid
            idx_selecao = 0

        colab_selecionado = st.selectbox(
            "Digite ou selecione o nome do colaborador:",
            options=opcoes_select,
            index=idx_selecao
        )

        # Se houver um colaborador válido selecionado, abre o formulário
        if colab_selecionado and colab_selecionado != "Selecione um colaborador...":
            dados_avulso = df[df["Nome"] == colab_selecionado].iloc[0]
            
            c_info, c_form = st.columns([1, 1])
            with c_info:
                st.info(
                    f"👤 **Colaborador:** {dados_avulso['Nome']}\n\n"
                    f"👔 **Cargo:** {dados_avulso.get('Cargo', 'Não informado')}\n\n"
                    f"🏭 **Setor:** {dados_avulso.get('Setor', 'Não informado')}\n\n"
                    f"🏢 **Unidade:** {aba_nome}"
                )
            
            with c_form:
                with st.expander("📄 Gerar Formulário de Agendamento", expanded=True):
                    tipo_avulso = st.selectbox("Tipo de Exame", ["PERIÓDICO", "MUDANÇA DE RISCO", "RETORNO", "ADMISSIONAL"], key="tipo_avulso")
                    dt_avulso = st.date_input("Data sugerida", value=hoje + timedelta(days=2), key="data_avulso")
                    btn_doc_avulso = gerar_docx(dados_avulso, tipo_avulso, dt_avulso, aba_nome)
                    st.download_button(
                        label="📥 Baixar Guia de Agendamento (DOCX)",
                        data=btn_doc_avulso,
                        file_name=f"ASO_{dados_avulso['Nome']}.docx",
                        key="btn_download_avulso"
                    )

        st.markdown("<hr style='margin: 25px 0;'>", unsafe_allow_html=True)

        # -------------------------------------------------------------
        # 2. MONITORAMENTO DE ALERTAS DE VENCIMENTO
        # -------------------------------------------------------------
        if "Venc" in df.columns:
            df["Venc"] = pd.to_datetime(df["Venc"], dayfirst=True, errors="coerce")
            df_alertas = df.dropna(subset=["Venc"]).copy()
            df_alertas["Dias"] = (df_alertas["Venc"] - hoje).dt.days
            alertas = df_alertas[df_alertas["Venc"] <= hoje + timedelta(days=10)].copy().sort_values(by="Venc")

            c1, c2, c3, c4 = st.columns(4)
            c1.metric("🏢 Unidade", aba_nome)
            c2.metric("👥 Colaboradores", len(df))
            c3.metric("⚠️ Pendentes", len(alertas))
            c4.metric("🚨 Vencidos", len(alertas[alertas["Dias"] < 0]))

            st.markdown("#### 📋 Alertas de Vencimento da Unidade")
            cols = st.columns(2)
            for idx, (_, row) in enumerate(alertas.iterrows()):
                col = cols[idx % 2]
                dias = int(row["Dias"])
                cor, fundo, status = ("#ef4444", "#fff1f2", "🚨 ASO VENCIDO") if dias < 0 else (("#f59e0b", "#fff7ed", f"⚠️ Vence em {dias} dias") if dias <= 3 else ("#10b981", "#ecfdf5", f"✅ Vence em {dias} dias"))

                html_card = f"""
                <div style="background:{fundo}; border-left:8px solid {cor}; border-radius:18px; padding:22px; margin-bottom:15px; box-shadow:0 4px 18px rgba(0,0,0,0.08); font-family:Arial;">
                    <div style="font-size:22px; font-weight:700; color:#111827; margin-bottom:10px;">{row['Nome']}</div>
                    <div style="color:#475569; font-size:15px; line-height:1.8;">👔 <b>Cargo:</b> {row['Cargo']}<br>🏭 <b>Setor:</b> {row['Setor']}<br>📅 <b>Vencimento:</b> {row['Venc'].strftime('%d/%m/%Y')}</div>
                    <div style="margin-top:15px; font-size:16px; font-weight:bold; color:{cor};">{status}</div>
                </div>"""

                with col:
                    components.html(html_card, height=250)
                    with st.expander(f"📄 Gerar Agendamento - {str(row['Nome']).split()[0]}"):
                        tipo = st.selectbox("Tipo de Exame", ["PERIÓDICO", "MUDANÇA DE RISCO", "RETORNO"], key=f"t_{idx}")
                        dt_s = st.date_input("Data sugerida", value=hoje + timedelta(days=2), key=f"d_{idx}")
                        btn_doc = gerar_docx(row, tipo, dt_s, aba_nome)
                        st.download_button(label="📥 Baixar Documento", data=btn_doc, file_name=f"ASO_{row['Nome']}.docx", key=f"b_{idx}")
except Exception as e:
    st.error(f"Erro: {e}")

st.markdown("""<div class="footer">© 2026 Gestão Documentos | Desenvolvido por: Dilceu Junior</div>""", unsafe_allow_html=True)
