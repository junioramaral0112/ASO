import concurrent.futures
from datetime import date, datetime
import io
import os
import sqlite3

from docx.shared import Mm
from docxtpl import DocxTemplate, InlineImage
import google.generativeai as genai
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# =========================================================
# CONFIGURAÇÃO GEMINI IA
# =========================================================
api_key = st.secrets.get("GEMINI_API_KEY", os.environ.get("GEMINI_API_KEY", ""))
model = None

if api_key:
    try:
        genai.configure(api_key=api_key)
        model = genai.GenerativeModel("gemini-1.5-flash")
    except Exception:
        model = None

PATH_BASE = os.path.dirname(__file__) if "__file__" in locals() else "."


def _chamar_api_gemini(prompt: str) -> str:
    if not model:
        raise RuntimeError("Modelo Gemini não configurado.")
    response = model.generate_content(prompt)
    return response.text.strip()


def gerar_texto_ia(pergunta: str, diagnostico: str = "") -> str:
    """Gera justificativa e ação corretiva com limite de 5s para não travar a aplicação."""
    texto_padrao = f"Risco identificado: {diagnostico} | Realizar adequação imediata conforme Norma Regulamentadora."
    if not api_key or not model:
        return texto_padrao

    prompt = (
        f"Você é um especialista SSMA. Analise o seguinte desvio de campo:\n"
        f"Item: {pergunta}\n"
        f"Diagnóstico: {diagnostico}\n"
        f"Responda EXATAMENTE no formato: JUSTIFICATIVA | AÇÃO CORRETIVA"
    )

    try:
        with concurrent.futures.ThreadPoolExecutor(max_workers=1) as executor:
            futuro = executor.submit(_chamar_api_gemini, prompt)
            return futuro.result(timeout=5)
    except Exception:
        return texto_padrao


# =========================================================
# BANCO DE DADOS
# =========================================================
DB_NAME = "jornada_ssma.db"


def init_db():
    with sqlite3.connect(DB_NAME) as conn:
        c = conn.cursor()
        c.execute(
            """
            CREATE TABLE IF NOT EXISTS inspecoes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                rel_num TEXT,
                unidade TEXT,
                data TEXT,
                auditor TEXT,
                arquivo_word BLOB
            )
        """
        )
        c.execute(
            """
            CREATE TABLE IF NOT EXISTS plano_acao (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                inspecao_id INTEGER,
                item TEXT,
                nr TEXT,
                obs TEXT,
                foto_antes BLOB,
                foto_depois BLOB,
                responsavel TEXT,
                prazo TEXT,
                status TEXT
            )
        """
        )
        conn.commit()


def get_next_rel_num() -> str:
    ano = datetime.now().year
    try:
        with sqlite3.connect(DB_NAME) as conn:
            c = conn.cursor()
            c.execute(
                "SELECT COUNT(*) FROM inspecoes WHERE rel_num LIKE ?",
                (f"%/{ano}",),
            )
            count = c.fetchone()[0]
            return f"{(count + 1):04d}/{ano}"
    except Exception:
        return f"0001/{ano}"


# =========================================================
# CONFIGURAÇÕES DE NEGÓCIO - FILIAIS E CHECKLIST
# =========================================================
UNIDADES = {
    "SÃO JOSÉ (MATRIZ)": {
        "CNPJ": "83.675.413/0001-01",
        "ENDERECO": "BR 101, km 210 / Bairro: Picadas do Sul – São José – SC",
    },
    "CHAPECÓ": {
        "CNPJ": "83.675.413/0002-84",
        "ENDERECO": "Rua Xanxerê, 360E – Bairro Líder – Chapecó – SC",
    },
    "CRICIÚMA": {
        "CNPJ": "83.675.413/0003-65",
        "ENDERECO": "Rodovia Jorge Lacerda, 4555 – Bairro Sangão – Criciúma – SC",
    },
    "LAGES": {
        "CNPJ": "83.675.413/0004-46",
        "ENDERECO": "Av. Dom Pedro II, 2887 – São Cristóvão – Lages – SC",
    },
    "CURITIBA": {
        "CNPJ": "83.675.413/0005-27",
        "ENDERECO": (
            "BR 116 (Linha Verde), 17350 – Pinheirinho – Curitiba – PR"
        ),
    },
    "CASCAVEL": {
        "CNPJ": "83.675.413/0006-08",
        "ENDERECO": "BR 277, Km 594 – Trevo Cataratas – Cascavel – PR",
    },
    "LONDRINA": {
        "CNPJ": "83.675.413/0007-90",
        "ENDERECO": "Av. Tiradentes, 5500 – Jardim Rosicler – Londrina – PR",
    },
    "PORTO ALEGRE": {
        "CNPJ": "83.675.413/0008-71",
        "ENDERECO": "Av. Severo Dullius, 1395 – São João – Porto Alegre – RS",
    },
    "PASSO FUNDO": {
        "CNPJ": "83.675.413/0009-52",
        "ENDERECO": (
            "RS 153, Km 02 – Nossa Senhora Aparecida – Passo Fundo – RS"
        ),
    },
    "CAXIAS DO SUL": {
        "CNPJ": "83.675.413/0010-94",
        "ENDERECO": "BR 116, Km 144 – São Ciro – Caxias do Sul – RS",
    },
    "JOINVILLE": {
        "CNPJ": "83.675.413/0011-75",
        "ENDERECO": "BR 101, KM 17 – Pirabeiraba – Joinville – SC",
    },
    "MARINGÁ": {
        "CNPJ": "83.675.413/0012-56",
        "ENDERECO": "PR-317, Km 05 – Parque Industrial – Maringá – PR",
    },
    "PELOTAS": {
        "CNPJ": "83.675.413/0013-37",
        "ENDERECO": "BR 116, Km 524 – Três Vendas – Pelotas – RS",
    },
    "SANTA MARIA": {
        "CNPJ": "83.675.413/0014-18",
        "ENDERECO": "BR 287, Km 242 – Camobi – Santa Maria – RS",
    },
}

TEMAS_CHECKLIST = {
    "Módulo Operacional": {
        "NR-12: Proteções em partes móveis/transmissão?": "NR 12",
        "NR-13: Vaso de pressão com teste hidrostático válido?": "NR 13",
        "Certificação: Elevação com selo de carga/revisão?": "NR 11",
        "Apoio: Uso de cavaletes após levantamento?": "NR 11",
        "Pneus: Gaiola de proteção para enchimento disponível?": "NR 12",
        "Bloqueio LOTO: Freio/travas em manutenções?": "NR 12",
        "Ferramentas Manuais: Livres de improvisos/gambiarras?": "NR 12",
        "Ferramentas Elétricas: Cabos/plugs íntegros?": "NR 10",
        "Espaçamento: Mínimo 60cm entre talhas e equipamentos?": "NR 11",
        "Documentação: I.T./LUP disponíveis no setor?": "Gestão",
    },
    "Módulo Ambiental": {
        "FDS (FISPQs): Disponíveis e atualizadas?": "NR 26",
        "Identificação: Químicos rotulados corretamente?": "NR 26",
        "Contenção: Kit de derramamento completo?": "NR 20",
        "Armazenamento: Inflamáveis longe de calor/faíscas?": "NR 20",
        "Resíduos: Descarte de óleos via logística reversa?": "CONAMA",
        "5S: Piso livre de óleo, graxa ou lixo industrial?": "NR 17",
        "Coleta Seletiva: Lixeiras identificadas e esvaziadas?": "Gestão",
        "Organização: Mezaninos sem sobrecarga?": "NR 11",
        "Higiene: Lava-olhos higienizado e funcional?": "NR 24",
        "EPI Químico: Proteção ocular/luvas disponíveis?": "NR 06",
    },
    "Módulo Infraestrutura": {
        "Treinamento: Executores com Integração em dia?": "NR 01",
        "NR-23: Extintores desobstruídos e com pressão?": "NR 23",
        "Emergência: Saídas e corredores livres?": "NR 23",
        "Quedas: Fossos/mezaninos com guarda-corpo?": "NR 12",
        "Escadas: Corrimão e fita antiderrapante?": "NR 12",
        "Elétrica: Painéis fechados e área de 1m livre?": "NR 10",
        "Iluminação: Adequada para a atividade?": "NR 17",
        "Ergonomia: Riscos de posturas/cargas?": "NR 17",
        "Primeiros Socorros: Maleta na validade?": "NR 07",
        "Acesso: Entrada de não autorizados controlada?": "Gestão",
    },
}


# =========================================================
# APLICAÇÃO PRINCIPAL
# =========================================================
def main():
    st.set_page_config(page_title="Gestão SSMA Macromaq", layout="wide")
    init_db()

    menu = [
        "📋 Nova Inspeção",
        "📊 Indicadores SSMA",
        "✅ Tratativas por Unidade",
        "📂 Histórico de Relatórios",
    ]

    choice = st.sidebar.selectbox("Navegação", menu)

    # -----------------------------------------------------
    # ABA 1 - NOVA INSPEÇÃO
    # -----------------------------------------------------
    if choice == "📋 Nova Inspeção":
        col_icon, col_text = st.columns([0.1, 0.9])
        with col_icon:
            icone_caminho = os.path.join(PATH_BASE, "Auditoria - ícon.png")
            if os.path.exists(icone_caminho):
                st.image(icone_caminho, width=70)
            else:
                st.markdown("## 📋")

        with col_text:
            st.title("Auditoria de Visita Técnica SSMA")

        with st.sidebar:
            u_sel = st.selectbox("Unidade", list(UNIDADES.keys()))
            auditor = st.selectbox(
                "Auditor",
                ["Jessica Vieira", "Dilceu Amaral Jr", "Simone Francisco"],
            )
            rel_num = st.text_input(
                "Nº Relatório", value=get_next_rel_num(), disabled=True
            )
            gerente = st.text_input("Gerente da Unidade", "Deveide Mafra")

        if "auditoria" not in st.session_state:
            st.session_state.auditoria = {
                p: {
                    "status": "Conforme",
                    "fotos": [],
                    "obs": "",
                    "prazo": date.today(),
                    "responsavel": "Guilherme Almeida",
                }
                for tema in TEMAS_CHECKLIST.values()
                for p in tema.keys()
            }

        if "docx_pronto" not in st.session_state:
            st.session_state.docx_pronto = None
        if "docx_nome" not in st.session_state:
            st.session_state.docx_nome = ""

        # Formulário do Checklist
        for tema, itens in TEMAS_CHECKLIST.items():
            with st.expander(f"📁 {tema}", expanded=True):
                for pergunta, nr in itens.items():
                    res = st.radio(
                        f"**{pergunta}**",
                        ["Conforme", "Não Conforme", "Não Aplica"],
                        horizontal=True,
                        key=f"r_{pergunta}",
                    )
                    st.session_state.auditoria[pergunta]["status"] = res

                    if res == "Não Conforme":
                        fotos = st.file_uploader(
                            "Fotos da Evidência",
                            type=["png", "jpg", "jpeg"],
                            accept_multiple_files=True,
                            key=f"up_{pergunta}",
                        )
                        if fotos:
                            st.session_state.auditoria[pergunta]["fotos"] = [
                                f.getvalue() for f in fotos
                            ]

                        st.session_state.auditoria[pergunta]["obs"] = (
                            st.text_area(
                                "Diagnóstico / Apontamento:",
                                key=f"obs_{pergunta}",
                            )
                        )
                        st.session_state.auditoria[pergunta]["responsavel"] = (
                            st.selectbox(
                                "Responsável:",
                                ["Guilherme Almeida", "Luiz", "Deveide Mafra"],
                                key=f"resp_{pergunta}",
                            )
                        )
                        st.session_state.auditoria[pergunta]["prazo"] = (
                            st.date_input(
                                "Prazo Limite:", key=f"praz_{pergunta}"
                            )
                        )
                    st.divider()

        # Botão de emissão
        if st.button("🚀 FINALIZAR E SALVAR RELATÓRIO", type="primary"):
            caminho_template = os.path.join(
                PATH_BASE, "JORNADA_SSMA_Template.docx"
            )
            if not os.path.exists(caminho_template):
                st.error(
                    f"Ficheiro de modelo não encontrado: '{caminho_template}'. Verifique a raiz do repositório."
                )
            else:
                with st.status(
                    "A processar e a gerar relatório...", expanded=True
                ) as status_box:
                    try:
                        status_box.write(
                            "Passo 1/3: A recolher respostas e a formatar itens..."
                        )
                        doc = DocxTemplate(caminho_template)
                        plano_acao = []
                        conformes = []
                        constatacoes = []

                        for t, its in TEMAS_CHECKLIST.items():
                            for p, nr in its.items():
                                v = st.session_state.auditoria[p]

                                if v["status"] == "Não Conforme":
                                    obs_limpa = (
                                        v["obs"].strip()
                                        if v["obs"]
                                        else "Não especificado"
                                    )

                                    if obs_limpa != "Não especificado":
                                        analise = gerar_texto_ia(p, obs_limpa)
                                    else:
                                        analise = "Risco identificado | Realizar adequação conforme NR"

                                    partes = analise.split("|")
                                    why = (
                                        partes[0].strip()
                                        if len(partes) > 0
                                        else "Risco identificado"
                                    )
                                    how = (
                                        partes[1].strip()
                                        if len(partes) > 1
                                        else "Adequar conforme NR"
                                    )

                                    imgs = [
                                        InlineImage(
                                            doc, io.BytesIO(fb), width=Mm(65)
                                        )
                                        for fb in v["fotos"]
                                        if fb
                                    ]

                                    constatacoes.append({
                                        "setor": p,
                                        "nr": nr,
                                        "descricao": obs_limpa,
                                        "fotos": imgs,
                                    })

                                    prazo_formatado = (
                                        v["prazo"].strftime("%d/%m/%Y")
                                        if isinstance(
                                            v["prazo"], (date, datetime)
                                        )
                                        else str(v["prazo"])
                                    )

                                    plano_acao.append({
                                        "pergunta": p,
                                        "why": why,
                                        "how": how,
                                        "responsavel": v["responsavel"],
                                        "prazo": prazo_formatado,
                                    })

                                elif v["status"] == "Conforme":
                                    conformes.append({
                                        "item": p,
                                        "obs": "Item verificado em conformidade com as diretrizes e requisitos de SSMA.",
                                    })

                        data_hoje = datetime.now().strftime("%d/%m/%Y")
                        contexto = {
                            "relatorio_num": rel_num,
                            "data_inspecao": data_hoje,
                            "unidade": u_sel,
                            "auditor_selecionado": auditor,
                            "gerente_unidade": gerente,
                            "empreendimento": f"Macromaq - {u_sel}",
                            "local": UNIDADES[u_sel]["ENDERECO"],
                            "cnpj": UNIDADES[u_sel]["CNPJ"],
                            "constatacoes": constatacoes,
                            "plano_acao": plano_acao,
                            "conformes": conformes,
                        }

                        status_box.write(
                            "Passo 2/3: A preencher ficheiro Word (.docx)..."
                        )
                        doc.render(contexto)
                        out = io.BytesIO()
                        doc.save(out)
                        out.seek(0)
                        word_binary = out.getvalue()

                        status_box.write(
                            "Passo 3/3: A gravar na base de dados SQLite..."
                        )
                        with sqlite3.connect(DB_NAME) as conn:
                            c = conn.cursor()
                            c.execute(
                                """
                                INSERT INTO inspecoes (rel_num, unidade, data, auditor, arquivo_word)
                                VALUES (?, ?, ?, ?, ?)
                                """,
                                (
                                    rel_num,
                                    u_sel,
                                    data_hoje,
                                    auditor,
                                    word_binary,
                                ),
                            )
                            ins_id = c.lastrowid

                            for item_p in plano_acao:
                                c.execute(
                                    """
                                    INSERT INTO plano_acao (inspecao_id, item, nr, obs, responsavel, prazo, status)
                                    VALUES (?, ?, ?, ?, ?, ?, ?)
                                    """,
                                    (
                                        ins_id,
                                        item_p["pergunta"],
                                        "NR",
                                        item_p["why"],
                                        item_p["responsavel"],
                                        item_p["prazo"],
                                        "Pendente",
                                    ),
                                )
                            conn.commit()

                        st.session_state.docx_pronto = word_binary
                        st.session_state.docx_nome = f"Relatorio_{u_sel}_{rel_num.replace('/', '-')}.docx"
                        status_box.update(
                            label="Relatório finalizado com sucesso!",
                            state="complete",
                            expanded=False,
                        )
                        st.success(
                            f"✅ Relatório {rel_num} pronto para descarregar!"
                        )

                    except Exception as e:
                        status_box.update(
                            label="Falha no processo!",
                            state="error",
                            expanded=True,
                        )
                        st.error(f"Erro detalhado: {e}")

        # Botão de download
        if st.session_state.docx_pronto is not None:
            st.download_button(
                label=(
                    f"📥 Descarregar Relatório Word"
                    f" ({st.session_state.docx_nome})"
                ),
                data=st.session_state.docx_pronto,
                file_name=st.session_state.docx_nome,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="primary",
            )

    # -----------------------------------------------------
    # ABA 2 - INDICADORES
    # -----------------------------------------------------
    elif choice == "📊 Indicadores SSMA":
        st.markdown(
            "<h1 style='text-align: center; color: #1c3d5a;'>Dashboard"
            " Executivo SSMA</h1>",
            unsafe_allow_html=True,
        )

        try:
            with sqlite3.connect(DB_NAME) as conn:
                df = pd.read_sql_query(
                    """
                    SELECT p.*, i.unidade, i.data as data_auditoria
                    FROM plano_acao p
                    JOIN inspecoes i ON p.inspecao_id = i.id
                    """,
                    conn,
                )
        except Exception:
            df = pd.DataFrame()

        if not df.empty:
            df["data_auditoria"] = pd.to_datetime(
                df["data_auditoria"], dayfirst=True
            )

            with st.sidebar:
                st.header("Filtros")
                u_f = st.selectbox(
                    "Unidade:", ["TODAS"] + list(UNIDADES.keys())
                )
                d_ini = st.date_input("Data Início:", df["data_auditoria"].min())
                d_fim = st.date_input("Data Fim:", date.today())

            mask = (df["data_auditoria"].dt.date >= d_ini) & (
                df["data_auditoria"].dt.date <= d_fim
            )
            if u_f != "TODAS":
                mask &= df["unidade"] == u_f
            df = df[mask]

            total_p = len(df)
            resolvidos = len(df[df["status"] == "Concluído"])
            eficacia = (resolvidos / total_p * 100) if total_p > 0 else 0

            c1, c2, c3 = st.columns(3)
            with c1:
                fig = go.Figure(
                    go.Indicator(
                        mode="gauge+number",
                        value=eficacia,
                        title={"text": "Eficácia (%)"},
                        gauge={"bar": {"color": "#2ecc71"}},
                    )
                )
                st.plotly_chart(fig, use_container_width=True)

            with c2:
                st.markdown(
                    f"""
                    <div style='background-color:white; padding:20px; border-radius:10px; border-left: 8px solid #e74c3c; text-align:center;'>
                        <h4 style='margin:0; color:#555;'>Pendências Abertas</h4>
                        <h1 style='margin:0; color:#e74c3c; font-size:48px;'>{total_p - resolvidos}</h1>
                    </div>
                """,
                    unsafe_allow_html=True,
                )

            with c3:
                st.markdown(
                    f"""
                    <div style='background-color:white; padding:20px; border-radius:10px; border-left: 8px solid #2ecc71; text-align:center;'>
                        <h4 style='margin:0; color:#555;'>Ações Concluídas</h4>
                        <h1 style='margin:0; color:#2ecc71; font-size:48px;'>{resolvidos}</h1>
                    </div>
                """,
                    unsafe_allow_html=True,
                )

            st.divider()
            ca, cb = st.columns([1.2, 0.8])

            with ca:
                st.write("### Ocorrências por Unidade")
                fig_bar = px.histogram(
                    df,
                    x="unidade",
                    color="status",
                    barmode="group",
                    color_discrete_map={
                        "Pendente": "#e74c3c",
                        "Concluído": "#2ecc71",
                    },
                )
                st.plotly_chart(fig_bar, use_container_width=True)

            with cb:
                st.write("### Distribuição dos Status")
                fig_p = px.pie(
                    df,
                    names="status",
                    hole=0.5,
                    color="status",
                    color_discrete_map={
                        "Pendente": "#e74c3c",
                        "Concluído": "#2ecc71",
                    },
                )
                st.plotly_chart(fig_p, use_container_width=True)
        else:
            st.info("Nenhuma inspeção registada de momento.")

    # -----------------------------------------------------
    # ABA 3 - TRATATIVAS
    # -----------------------------------------------------
    elif choice == "✅ Tratativas por Unidade":
        col_f1, col_f2 = st.columns([0.1, 0.9])
        with col_f1:
            icone_follow = os.path.join(PATH_BASE, "Follow UP - Icon.png")
            if os.path.exists(icone_follow):
                st.image(icone_follow, width=70)
            else:
                st.markdown("## 🔄")
        with col_f2:
            st.title("Follow-up de Pendências")

        try:
            with sqlite3.connect(DB_NAME) as conn:
                df_p = pd.read_sql_query(
                    """
                    SELECT p.*, i.unidade
                    FROM plano_acao p
                    JOIN inspecoes i ON p.inspecao_id = i.id
                    WHERE p.status = 'Pendente'
                    """,
                    conn,
                )

                if not df_p.empty:
                    unidades_disponiveis = df_p["unidade"].unique()
                    u_f = st.selectbox(
                        "Selecione a Unidade:", unidades_disponiveis
                    )
                    itens_filtrados = df_p[df_p["unidade"] == u_f]

                    for _, row in itens_filtrados.iterrows():
                        with st.expander(
                            f"📌 {row['item']} | Prazo: {row['prazo']}"
                        ):
                            st.error(f"**Diagnóstico:** {row['obs']}")
                            st.write(f"**Responsável:** {row['responsavel']}")
                            f_dep = st.file_uploader(
                                "Anexar Comprovativo / Foto de Baixa"
                                f" #{row['id']}",
                                key=f"f_{row['id']}",
                            )

                            if st.button(
                                f"Concluir Pendência #{row['id']}",
                                type="primary",
                                key=f"btn_{row['id']}",
                            ):
                                foto_bytes = f_dep.getvalue() if f_dep else None
                                cursor = conn.cursor()
                                cursor.execute(
                                    """
                                    UPDATE plano_acao
                                    SET status='Concluído', foto_depois=?
                                    WHERE id=?
                                    """,
                                    (foto_bytes, row["id"]),
                                )
                                conn.commit()
                                st.success(
                                    f"Item #{row['id']} concluído com sucesso!"
                                )
                                st.rerun()
                else:
                    st.success(
                        "🎉 Todas as unidades estão 100% em conformidade!"
                        " Nenhuma pendência aberta."
                    )
        except Exception as e:
            st.error(f"Erro ao carregar pendências: {e}")

    # -----------------------------------------------------
    # ABA 4 - HISTÓRICO
    # -----------------------------------------------------
    elif choice == "📂 Histórico de Relatórios":
        col_h1, col_h2 = st.columns([0.1, 0.9])
        with col_h1:
            icone_hist = os.path.join(PATH_BASE, "Auditoria - ícon.png")
            if os.path.exists(icone_hist):
                st.image(icone_hist, width=70)
            else:
                st.markdown("## 📂")
        with col_h2:
            st.title("Histórico de Relatórios")

        try:
            with sqlite3.connect(DB_NAME) as conn:
                df_hist = pd.read_sql_query(
                    """
                    SELECT id, rel_num, unidade, data, auditor
                    FROM inspecoes
                    ORDER BY id DESC
                    """,
                    conn,
                )

                if not df_hist.empty:
                    for _, row in df_hist.iterrows():
                        c1, c2, c3, c4 = st.columns([2, 2, 2, 2])
                        c1.write(f"**Nº:** {row['rel_num']}")
                        c2.write(f"**Unidade:** {row['unidade']}")
                        c3.write(f"**Data:** {row['data']}")

                        cursor = conn.cursor()
                        cursor.execute(
                            "SELECT arquivo_word FROM inspecoes WHERE id=?",
                            (row["id"],),
                        )
                        dado = cursor.fetchone()
                        blob = dado[0] if dado else None

                        if blob:
                            c4.download_button(
                                "📥 Descarregar DOCX",
                                blob,
                                f"Relatorio_{row['rel_num'].replace('/', '-')}.docx",
                                key=f"h_{row['id']}",
                            )
                        st.divider()
                else:
                    st.info("Nenhum relatório emitido até ao momento.")
        except Exception:
            st.info("Histórico ainda vazio.")


if __name__ == "__main__":
    main()
