import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
from datetime import datetime, timedelta, time
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from io import BytesIO
import os
import urllib.request
import urllib.error
from PIL import Image
import base64
import unicodedata
from urllib.parse import unquote
import json
import html

# =========================================================
# CONFIGURAÇÃO DE CAMINHOS DINÂMICOS
# =========================================================

BASE_DIR = os.path.dirname(__file__) if "__file__" in locals() else "."

def localizar_arquivo(caminho_local, nome_arquivo):
    if os.path.exists(caminho_local):
        return caminho_local
    caminho_projeto = os.path.join(BASE_DIR, nome_arquivo)
    return caminho_projeto if os.path.exists(caminho_projeto) else None

PATH_ASO_IMG = localizar_arquivo(
    r"C:\Users\dilceu.gomes\Desktop\sistema_aso\ASO.png", "ASO.png"
)
PATH_LOGO = localizar_arquivo(
    r"C:\Users\dilceu.gomes\Desktop\sistema_aso\logo.png", "logo.png"
)
PATH_LOGO_DOC = localizar_arquivo(
    r"C:\Users\dilceu.gomes\Desktop\sistema_aso\adivitta.png", "adivitta.png"
)


def carregar_imagem_base64(path):
    if path and os.path.exists(path):
        with open(path, "rb") as img_file:
            return base64.b64encode(img_file.read()).decode()
    return None


def normalizar_texto(texto):
    if not isinstance(texto, str):
        return ""
    texto = (
        unicodedata.normalize("NFKD", texto)
        .encode("ASCII", "ignore")
        .decode("ASCII")
    )
    return texto.strip().upper()


def valor_texto(valor, padrao=""):
    """Converte qualquer valor vindo do pandas/Sheets em texto seguro."""
    if valor is None:
        return padrao
    try:
        if pd.isna(valor):
            return padrao
    except (TypeError, ValueError):
        pass
    texto = str(valor).strip()
    if texto.lower() in {"nan", "none", "nat"}:
        return padrao
    return texto


# =========================================================
# CONFIGURAÇÃO DA PÁGINA
# =========================================================

try:
    if PATH_ASO_IMG:
        page_icon = Image.open(PATH_ASO_IMG)
    else:
        page_icon = "🛡️"
except Exception:
    page_icon = "🛡️"

st.set_page_config(
    page_title="Gestão ASO Pro",
    layout="wide",
    page_icon=page_icon,
)

# =========================================================
# CSS GLOBAL E TOPO
# =========================================================

aso_base64 = carregar_imagem_base64(PATH_ASO_IMG)

st.markdown(
    f"""
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
""",
    unsafe_allow_html=True,
)

# =========================================================
# LÓGICA DE DADOS
# =========================================================

SHEET_ID = "1G_oVT9gK-n_jGh5R4g65qUwK_MfQGvCX-SA4NHNNflU"

# Mantidos exatamente conforme o código fornecido.
UNIDADES = {
    "C-CTBA": "145843404",
    "C-CHP": "1549718037",
    "S-SJ": "1712391604",
    "D-MG": "1071212860",
    "D-ITU": "1323067532",
}


def limpar_dataframe(df):
    """Limpeza defensiva sem perder a numeração real das linhas da planilha."""
    if df is None or df.empty:
        return pd.DataFrame()

    df = df.copy()
    df.columns = df.columns.astype(str).str.strip()

    # A exportação CSV começa com o cabeçalho na linha 1 da planilha.
    # Portanto, o primeiro registro de dados corresponde à linha 2.
    df["_SHEET_ROW"] = range(2, len(df) + 2)

    if "Nome" in df.columns:
        df["Nome"] = df["Nome"].apply(valor_texto)
        df = df[
            (df["Nome"] != "")
            & (~df["Nome"].str.lower().isin(["nan", "none"]))
        ].copy()

    # Campos textuais usados pelo aplicativo.
    for col_name in ["Cargo", "Setor", "UNIDADE", "Agendamento"]:
        if col_name in df.columns:
            df[col_name] = df[col_name].apply(valor_texto)

    return df


@st.cache_data(ttl=60)
def load_data(gid):
    url = (
        f"https://docs.google.com/spreadsheets/d/{SHEET_ID}"
        f"/export?format=csv&gid={str(gid).strip()}"
    )

    try:
        req = urllib.request.Request(
            url,
            headers={"User-Agent": "Mozilla/5.0"},
        )
        with urllib.request.urlopen(req, timeout=20) as resp:
            df = pd.read_csv(resp)

        return limpar_dataframe(df)

    except urllib.error.HTTPError as e:
        if e.code == 400:
            st.error(
                f"❌ O GID `{gid}` não pôde ser carregado pelo Google Sheets. "
                "Confira se a aba existe e se o GID da URL está correto."
            )
        else:
            st.error(f"Erro HTTP {e.code} ao baixar dados da planilha.")
        return pd.DataFrame()

    except Exception as ex:
        st.warning(f"Erro ao ler GID {gid}: {ex}")
        return pd.DataFrame()


@st.cache_data(ttl=60)
def carregar_base_completa():
    frames = []

    for u_nome, gid in UNIDADES.items():
        d = load_data(gid)

        if not d.empty and "Nome" in d.columns:
            d["UNIDADE_REAL"] = u_nome
            frames.append(d)

    if frames:
        return pd.concat(frames, ignore_index=True)

    return pd.DataFrame()


# =========================================================
# GOOGLE SHEETS — ESCRITA DO AGENDAMENTO
# =========================================================

def obter_google_client():
    """
    Cria cliente gspread usando Streamlit Secrets.

    Formato recomendado:
    [gcp_service_account]
    type = "service_account"
    project_id = "..."
    private_key_id = "..."
    private_key = "-----BEGIN PRIVATE KEY-----\\n...\\n-----END PRIVATE KEY-----\\n"
    client_email = "..."
    client_id = "..."
    auth_uri = "https://accounts.google.com/o/oauth2/auth"
    token_uri = "https://oauth2.googleapis.com/token"
    auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
    client_x509_cert_url = "..."
    """
    try:
        import gspread
    except ImportError as exc:
        raise RuntimeError(
            "A biblioteca 'gspread' não está instalada. "
            "Adicione gspread ao requirements.txt."
        ) from exc

    if "gcp_service_account" in st.secrets:
        info = dict(st.secrets["gcp_service_account"])

        # Alguns ambientes armazenam a chave com \\n literal.
        if "private_key" in info:
            info["private_key"] = str(info["private_key"]).replace("\\n", "\n")

        return gspread.service_account_from_dict(info)

    # Compatibilidade opcional com um Secret contendo o JSON inteiro.
    if "GOOGLE_SERVICE_ACCOUNT_JSON" in st.secrets:
        raw = st.secrets["GOOGLE_SERVICE_ACCOUNT_JSON"]

        try:
            info = json.loads(str(raw))
        except json.JSONDecodeError as exc:
            raise RuntimeError(
                "GOOGLE_SERVICE_ACCOUNT_JSON existe, mas não contém JSON válido."
            ) from exc

        if "private_key" in info:
            info["private_key"] = str(info["private_key"]).replace("\\n", "\n")

        return gspread.service_account_from_dict(info)

    raise RuntimeError(
        "Credenciais Google não encontradas nos Secrets. "
        "Configure [gcp_service_account] no Streamlit Secrets."
    )


def obter_worksheet_por_gid(gspread_client, gid):
    """Localiza a aba pelo GID real, sem depender do nome da aba."""
    planilha = gspread_client.open_by_key(SHEET_ID)

    gid_str = str(gid).strip()

    for worksheet in planilha.worksheets():
        if str(worksheet.id) == gid_str:
            return worksheet

    raise RuntimeError(
        f"A aba com GID {gid_str} não foi encontrada na planilha."
    )


def localizar_coluna_agendamento(worksheet):
    """Localiza a coluna pelo nome do cabeçalho, nunca por posição fixa."""
    headers = worksheet.row_values(1)

    for index, header in enumerate(headers, start=1):
        if normalizar_texto(header) == "AGENDAMENTO":
            return index

    raise RuntimeError(
        f'A coluna "Agendamento" não foi encontrada na aba "{worksheet.title}". '
        "Nenhuma coluna nova foi criada automaticamente."
    )


def gravar_agendamento(gid, sheet_row, texto_agendamento):
    """
    Atualiza SOMENTE a célula da coluna Agendamento na linha informada.
    Não sobrescreve a linha inteira.
    """
    try:
        cliente = obter_google_client()
        worksheet = obter_worksheet_por_gid(cliente, gid)
        coluna = localizar_coluna_agendamento(worksheet)

        # gspread usa índice de linha/coluna começando em 1.
        worksheet.update_cell(
            int(sheet_row),
            int(coluna),
            texto_agendamento,
        )

        # Limpa os caches para a nova informação aparecer imediatamente.
        try:
            load_data.clear()
        except Exception:
            pass

        try:
            carregar_base_completa.clear()
        except Exception:
            pass

        return True, ""

    except Exception as exc:
        return False, str(exc)


def salvar_agendamento(gid, sheet_row, data_agendamento, hora_agendamento, observacao):
    """Monta o texto e grava o agendamento."""
    data_str = data_agendamento.strftime("%d/%m/%Y")
    hora_str = hora_agendamento.strftime("%H:%M")

    obs = valor_texto(observacao)

    texto = f"{data_str} às {hora_str}"

    if obs:
        texto += f" — {obs}"

    return gravar_agendamento(gid, sheet_row, texto)


def limpar_agendamento(gid, sheet_row):
    """Limpa somente a célula Agendamento."""
    return gravar_agendamento(gid, sheet_row, "")


# =========================================================
# DOCUMENTO WORD
# =========================================================

def gerar_docx(dados, tipo, data_sugestao, unidade_padrao="MACROMAQ"):
    doc = Document()

    if PATH_LOGO_DOC and os.path.exists(PATH_LOGO_DOC):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run().add_picture(PATH_LOGO_DOC, width=Inches(1.2))

    titulo = doc.add_paragraph()
    titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = titulo.add_run(f"FORMULÁRIO PARA AGENDAMENTO {tipo}")
    run.bold = True
    run.font.size = Pt(16)

    table = doc.add_table(rows=7, cols=2)
    table.style = "Table Grid"

    labels = [
        "Nome Completo",
        "Cargo",
        "Setor",
        "Unidade",
        "Cliente",
        "Local",
        "Data Sugestão",
    ]

    nome_val = valor_texto(dados.get("Nome", ""))
    cargo_val = valor_texto(dados.get("Cargo", ""))
    setor_val = valor_texto(dados.get("Setor", ""))
    unid_val = valor_texto(
        dados.get("UNIDADE", dados.get("UNIDADE_REAL", unidade_padrao))
    )

    valores = [
        nome_val,
        cargo_val,
        setor_val,
        unid_val,
        "MACROMAQ",
        "Arapoti",
        data_sugestao.strftime("%d/%m/%Y"),
    ]

    for i, (label, value) in enumerate(zip(labels, valores)):
        table.rows[i].cells[0].text = label
        table.rows[i].cells[1].text = value

        if table.rows[i].cells[0].paragraphs[0].runs:
            table.rows[i].cells[0].paragraphs[0].runs[0].bold = True

    target = BytesIO()
    doc.save(target)
    return target.getvalue()


# =========================================================
# COMPONENTE DE AGENDAMENTO
# =========================================================

def renderizar_agendamento(row, gid, aba_nome, idx):
    """
    Interface de agendamento integrada ao card.
    Usa a linha real da planilha armazenada em _SHEET_ROW.
    """
    nome = valor_texto(row.get("Nome", "Colaborador"))
    sheet_row = row.get("_SHEET_ROW")

    if pd.isna(sheet_row) if not isinstance(sheet_row, str) else False:
        sheet_row = None

    if sheet_row is None:
        st.error("Não foi possível identificar a linha do colaborador na planilha.")
        return

    try:
        sheet_row = int(float(sheet_row))
    except (TypeError, ValueError):
        st.error("Não foi possível identificar a linha do colaborador na planilha.")
        return

    info_agendamento = valor_texto(row.get("Agendamento", ""))

    st.markdown("##### 📅 Agendamento")

    if info_agendamento:
        st.success(f"📌 AGENDADO — {info_agendamento}")

        modo = st.radio(
            "Ação",
            ["Manter agendamento", "Editar agendamento", "Desmarcar agendamento"],
            horizontal=True,
            key=f"ag_modo_{aba_nome}_{sheet_row}_{idx}",
        )

        if modo == "Editar agendamento":
            data_padrao = datetime.now().date()
            hora_padrao = time(8, 0)
            observacao_padrao = ""

            # Não tenta interpretar formatos antigos automaticamente.
            # O usuário pode informar os novos dados e substituir o texto atual.
            data_ag = st.date_input(
                "Data do agendamento",
                value=data_padrao,
                key=f"ag_data_edit_{aba_nome}_{sheet_row}_{idx}",
            )
            hora_ag = st.time_input(
                "Hora do agendamento",
                value=hora_padrao,
                key=f"ag_hora_edit_{aba_nome}_{sheet_row}_{idx}",
            )
            obs = st.text_area(
                "Observação",
                value=observacao_padrao,
                key=f"ag_obs_edit_{aba_nome}_{sheet_row}_{idx}",
                placeholder="Digite uma observação, se necessário.",
            )

            if st.button(
                "💾 Salvar alteração",
                key=f"ag_salvar_edit_{aba_nome}_{sheet_row}_{idx}",
                type="primary",
            ):
                ok, erro = salvar_agendamento(
                    gid,
                    sheet_row,
                    data_ag,
                    hora_ag,
                    obs,
                )

                if ok:
                    st.success("✅ Agendamento atualizado com sucesso.")
                    st.rerun()
                else:
                    st.error(
                        "❌ Não foi possível atualizar o agendamento.\n\n"
                        f"Detalhes: {erro}"
                    )

        elif modo == "Desmarcar agendamento":
            confirmar = st.checkbox(
                "Confirmo que quero remover este agendamento.",
                key=f"ag_confirma_remove_{aba_nome}_{sheet_row}_{idx}",
            )

            if st.button(
                "🗑️ Desmarcar agendamento",
                key=f"ag_remover_{aba_nome}_{sheet_row}_{idx}",
                disabled=not confirmar,
            ):
                ok, erro = limpar_agendamento(gid, sheet_row)

                if ok:
                    st.success("✅ Agendamento removido com sucesso.")
                    st.rerun()
                else:
                    st.error(
                        "❌ Não foi possível remover o agendamento.\n\n"
                        f"Detalhes: {erro}"
                    )

    else:
        marcar = st.checkbox(
            "Marcar como agendado",
            key=f"ag_marcar_{aba_nome}_{sheet_row}_{idx}",
        )

        if marcar:
            col_data, col_hora = st.columns(2)

            with col_data:
                data_ag = st.date_input(
                    "Data do agendamento",
                    value=datetime.now().date(),
                    key=f"ag_data_{aba_nome}_{sheet_row}_{idx}",
                )

            with col_hora:
                hora_ag = st.time_input(
                    "Hora do agendamento",
                    value=time(8, 0),
                    key=f"ag_hora_{aba_nome}_{sheet_row}_{idx}",
                )

            obs = st.text_area(
                "Observação",
                key=f"ag_obs_{aba_nome}_{sheet_row}_{idx}",
                placeholder="Ex.: Agendamento confirmado com a unidade.",
            )

            if st.button(
                "💾 Salvar agendamento",
                key=f"ag_salvar_{aba_nome}_{sheet_row}_{idx}",
                type="primary",
            ):
                ok, erro = salvar_agendamento(
                    gid,
                    sheet_row,
                    data_ag,
                    hora_ag,
                    obs,
                )

                if ok:
                    st.success("✅ Agendamento salvo com sucesso.")
                    st.rerun()
                else:
                    st.error(
                        "❌ Não foi possível salvar o agendamento.\n\n"
                        f"Detalhes: {erro}"
                    )


# =========================================================
# SIDEBAR
# =========================================================

if PATH_LOGO and os.path.exists(PATH_LOGO):
    st.sidebar.image(PATH_LOGO, width=300)

aba_nome = st.sidebar.selectbox(
    "🏢 Selecione a unidade",
    list(UNIDADES.keys()),
)

st.sidebar.success("✅ Sistema Online")
st.sidebar.markdown("<br>", unsafe_allow_html=True)
st.sidebar.markdown(
    "<div class='menu-titulo'>⚡ ACESSO RÁPIDO</div>",
    unsafe_allow_html=True,
)

if st.sidebar.button("📋 CHECKLIST SSMA"):
    st.switch_page("pages/app_ssma_ia.py")


# =========================================================
# BUSCA ATIVA E EMISSÃO DE GUIA AVULSA
# =========================================================

st.markdown("### 🔍 Busca Ativa e Emissão de Guia Avulsa")

df_todos = carregar_base_completa()

if not df_todos.empty and "Nome" in df_todos.columns:
    colab_param = st.query_params.get("colaborador", None)

    lista_todos_nomes = sorted(
        [
            valor_texto(n)
            for n in df_todos["Nome"].dropna().unique()
            if valor_texto(n)
        ]
    )

    nome_pre_selecionado = None

    if colab_param:
        nome_param_limpo = normalizar_texto(unquote(str(colab_param)))

        for n in lista_todos_nomes:
            n_normalizado = normalizar_texto(n)

            if (
                n_normalizado == nome_param_limpo
                or nome_param_limpo in n_normalizado
            ):
                nome_pre_selecionado = n
                break

        if not nome_pre_selecionado:
            nome_pre_selecionado = unquote(str(colab_param)).strip()

    if nome_pre_selecionado:
        if nome_pre_selecionado in lista_todos_nomes:
            opcoes_finais = [nome_pre_selecionado] + [
                x for x in lista_todos_nomes if x != nome_pre_selecionado
            ]
        else:
            opcoes_finais = [nome_pre_selecionado] + lista_todos_nomes

        idx_padrao = 0

    else:
        opcoes_finais = ["Selecione um colaborador..."] + lista_todos_nomes
        idx_padrao = 0

    colab_escolhido = st.selectbox(
        "Digite ou selecione o nome do colaborador:",
        options=opcoes_finais,
        index=idx_padrao,
    )

    if (
        colab_escolhido
        and colab_escolhido != "Selecione um colaborador..."
    ):
        registro_filtrado = df_todos[
            df_todos["Nome"].astype(str) == str(colab_escolhido)
        ]

        if not registro_filtrado.empty:
            dados_colab = registro_filtrado.iloc[0]
        else:
            dados_colab = {
                "Nome": str(colab_escolhido),
                "Cargo": "Geral",
                "Setor": "Operacional",
                "UNIDADE_REAL": "MACROMAQ",
            }

        partes = valor_texto(colab_escolhido).split()
        primeiro_nome = partes[0] if partes else "Colaborador"

        hoje = datetime.now()
        doc_bytes = gerar_docx(
            dados_colab,
            "PERIÓDICO",
            hoje + timedelta(days=2),
            aba_nome,
        )

        st.download_button(
            label=f"📥 Baixar Formulário de {primeiro_nome}",
            data=doc_bytes,
            file_name=f"ASO_{colab_escolhido}.docx",
            key="btn_download_direto",
        )

st.markdown("<hr style='margin: 25px 0;'>", unsafe_allow_html=True)


# =========================================================
# MONITORAMENTO DE ALERTAS DE VENCIMENTO
# =========================================================

try:
    gid_unidade = UNIDADES[aba_nome]
    df_unidade = load_data(gid_unidade)

    if (
        not df_unidade.empty
        and "Nome" in df_unidade.columns
        and "Venc" in df_unidade.columns
    ):
        hoje = datetime.now()

        df_unidade["Venc"] = pd.to_datetime(
            df_unidade["Venc"],
            dayfirst=True,
            errors="coerce",
        )

        df_alertas = df_unidade.dropna(
            subset=["Venc", "Nome"]
        ).copy()

        # Tratamento defensivo do Nome.
        df_alertas["Nome_Str"] = df_alertas["Nome"].apply(valor_texto)

        df_alertas = df_alertas[
            (df_alertas["Nome_Str"] != "")
            & (~df_alertas["Nome_Str"].str.lower().isin(["nan", "none"]))
        ].copy()

        # Elimina datas antigas/fantasmas.
        df_alertas = df_alertas[
            df_alertas["Venc"].dt.year >= 2020
        ].copy()

        df_alertas["Dias"] = (
            df_alertas["Venc"] - hoje
        ).dt.days

        alertas = (
            df_alertas[
                df_alertas["Venc"] <= hoje + timedelta(days=10)
            ]
            .copy()
            .sort_values(by="Venc")
        )

        c1, c2, c3, c4 = st.columns(4)

        c1.metric("🏢 Unidade", aba_nome)
        c2.metric("👥 Colaboradores", len(df_unidade))
        c3.metric("⚠️ Pendentes", len(alertas))
        c4.metric(
            "🚨 Vencidos",
            len(alertas[alertas["Dias"] < 0]),
        )

        st.markdown("#### 📋 Alertas de Vencimento da Unidade")

        if alertas.empty:
            st.info(
                "Nenhum ASO vencido ou a vencer nos próximos "
                "10 dias para esta unidade."
            )

        else:
            cols = st.columns(2)

            for idx, (_, row) in enumerate(alertas.iterrows()):
                col = cols[idx % 2]

                dias = int(row["Dias"])

                if dias < 0:
                    cor = "#ef4444"
                    fundo = "#fff1f2"
                    status = "🚨 ASO VENCIDO"
                elif dias <= 3:
                    cor = "#f59e0b"
                    fundo = "#fff7ed"
                    status = f"⚠️ Vence em {dias} dias"
                else:
                    cor = "#10b981"
                    fundo = "#ecfdf5"
                    status = f"✅ Vence em {dias} dias"

                nome_str = valor_texto(row.get("Nome_Str", "Colaborador"))

                cargo_str = valor_texto(
                    row.get("Cargo", ""),
                    "Não informado",
                )

                if not cargo_str:
                    cargo_str = "Não informado"

                setor_str = valor_texto(
                    row.get("Setor", ""),
                    "Não informado",
                )

                if not setor_str:
                    setor_str = "Não informado"

                venc_str = row["Venc"].strftime("%d/%m/%Y")

                # =================================================
                # LEITURA DO AGENDAMENTO
                # =================================================

                info_agendamento = valor_texto(
                    row.get("Agendamento", "")
                )

                badge_html = ""

                detalhes_agendamento_html = ""

                if info_agendamento:
                    badge_html = (
                        '<span style="background:#2563eb; color:white; '
                        'padding:4px 10px; border-radius:8px; '
                        'font-size:12px; font-weight:bold; '
                        'margin-left:10px;">📌 AGENDADO</span>'
                    )

                    detalhes_agendamento_html = f"""
                    <div style="margin-top:10px; padding:10px 14px;
                                background:#e0f2fe; border-radius:10px;
                                color:#0369a1; font-size:13px;
                                border:1px solid #bae6fd;">
                        🗓️ <b>Agendamento:</b>
                        {html.escape(info_agendamento)}
                    </div>
                    """

                html_card = f"""
                <div style="background:{fundo};
                            border-left:8px solid {cor};
                            border-radius:18px;
                            padding:22px;
                            margin-bottom:15px;
                            box-shadow:0 4px 18px rgba(0,0,0,0.08);
                            font-family:Arial;">

                    <div style="font-size:22px;
                                font-weight:700;
                                color:#111827;
                                margin-bottom:10px;
                                display:flex;
                                align-items:center;
                                flex-wrap:wrap;">
                        {html.escape(nome_str)} {badge_html}
                    </div>

                    <div style="color:#475569;
                                font-size:15px;
                                line-height:1.8;">
                        👔 <b>Cargo:</b> {html.escape(cargo_str)}<br>
                        🏭 <b>Setor:</b> {html.escape(setor_str)}<br>
                        📅 <b>Vencimento:</b> {venc_str}
                    </div>

                    {detalhes_agendamento_html}

                    <div style="margin-top:15px;
                                font-size:16px;
                                font-weight:bold;
                                color:{cor};">
                        {status}
                    </div>
                </div>
                """

                with col:
                    altura_card = 310 if info_agendamento else 240
                    components.html(
                        html_card,
                        height=altura_card,
                    )

                    partes_nome = nome_str.split()
                    primeiro_nome_card = (
                        partes_nome[0]
                        if partes_nome
                        else "Colaborador"
                    )

                    # =================================================
                    # AGENDAMENTO PERSISTENTE
                    # =================================================

                    with st.expander(
                        f"📅 Agendamento - {primeiro_nome_card}",
                        expanded=False,
                    ):
                        renderizar_agendamento(
                            row,
                            gid_unidade,
                            aba_nome,
                            idx,
                        )

                    # =================================================
                    # DOCUMENTO WORD
                    # =================================================

                    with st.expander(
                        f"📄 Formulário - {primeiro_nome_card}"
                    ):
                        tipo = st.selectbox(
                            "Tipo de Exame",
                            [
                                "PERIÓDICO",
                                "MUDANÇA DE RISCO",
                                "RETORNO",
                            ],
                            key=f"t_{aba_nome}_{idx}",
                        )

                        dt_s = st.date_input(
                            "Data sugerida para formulário",
                            value=hoje.date() + timedelta(days=2),
                            key=f"d_{aba_nome}_{idx}",
                        )

                        btn_doc = gerar_docx(
                            row,
                            tipo,
                            dt_s,
                            aba_nome,
                        )

                        st.download_button(
                            label="📥 Baixar Documento Word",
                            data=btn_doc,
                            file_name=f"ASO_{nome_str}.docx",
                            key=f"b_{aba_nome}_{idx}",
                        )

    else:
        st.warning(
            f"Sem dados carregados para a unidade {aba_nome}."
        )

except Exception as e:
    st.error(f"Erro ao processar dados: {e}")


st.markdown(
    """<div class="footer">
    © 2026 Gestão Documentos | Desenvolvido por: Dilceu Junior
    </div>""",
    unsafe_allow_html=True,
)
