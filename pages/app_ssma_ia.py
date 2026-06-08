import streamlit as st
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
from datetime import datetime, date
import io
import sqlite3
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import google.generativeai as genai
import os

# =========================================================
# CONFIGURAÇÃO GEMINI IA (ADAPTADO PARA O GITHUB/STREAMLIT)
# =========================================================
# Buscando a chave de forma segura através dos Secrets do Streamlit
if "GEMINI_API_KEY" in st.secrets:
    genai.configure(api_key=st.secrets["GEMINI_API_KEY"])
else:
    # Fallback caso você ainda esteja testando localmente sem secrets
    genai.configure(api_key="SUA_CHAVE_LOCAL_SE_NECESSARIO")

model = genai.GenerativeModel('gemini-1.5-flash')

# ADAPTAÇÃO GITHUB: Removido o caminho fixo do C: 
# Agora o sistema procura as imagens na mesma pasta raiz do projeto.
PATH_BASE = "" 

def gerar_texto_ia(pergunta, status, diagnostico=""):
    try:
        if status == "Conforme":
            prompt = f"Gere uma frase curta técnica confirmando que o item '{pergunta}' cumpre os requisitos SSMA da Macromaq."
        else:
            prompt = f"Especialista SSMA. Item: {pergunta}. Diagnóstico: {diagnostico}. Formato: JUSTIFICATIVA | AÇÃO CORRETIVA."
        response = model.generate_content(prompt)
        return response.text.strip()
    except:
        return "Item em conformidade." if status == "Conforme" else "Risco identificado | Realizar adequação conforme NR."

# =========================================================
# BANCO DE DADOS (ESTRUTURA COMPLETA)
# =========================================================
def init_db():
    conn = sqlite3.connect('jornada_ssma.db')
    c = conn.cursor()

    c.execute('''
        CREATE TABLE IF NOT EXISTS inspecoes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            rel_num TEXT,
            unidade TEXT,
            data TEXT,
            auditor TEXT,
            arquivo_word BLOB
        )
    ''')

    c.execute('''
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
    ''')

    conn.commit()
    conn.close()

def get_next_rel_num():
    ano = datetime.now().year
    try:
        conn = sqlite3.connect('jornada_ssma.db')
        c = conn.cursor()
        c.execute(
            "SELECT COUNT(*) FROM inspecoes WHERE rel_num LIKE ?",
            (f'%/{ano}',)
        )
        count = c.fetchone()[0]
        conn.close()
        return f"{(count + 1):04d}/{ano}"
    except:
        return f"0001/{ano}"

# =========================================================
# CONFIGURAÇÕES DE NEGÓCIO
# =========================================================
UNIDADES = {
    "SÃO JOSÉ": {
        "CNPJ": "83.675.413/0001-01",
        "ENDERECO": "BR 101, km 210 / Bairro: Picadas do Sul – São José – SC"
    },
    "CHAPECÓ": {
        "CNPJ": "83.675.413/0002-84",
        "ENDERECO": "Rua Xanxerê, 360E – Bairro Líder – Chapecó/SC"
    },
    "JOINVILLE": {
        "CNPJ": "83.675.413/0011-75",
        "ENDERECO": "BR101, KM17 – Joinville / SC"
    }
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
        "Documentação: I.T./LUP disponíveis no setor?": "Gestão"
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
        "Higiene: Lava-olhos higienizado e functional?": "NR 24",
        "EPI Químico: Proteção ocular/luvas disponíveis?": "NR 06"
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
        "Acesso: Entrada de não autorizados controlada?": "Gestão"
    }
}

# =========================================================
# APP PRINCIPAL
# =========================================================
def main():
    st.set_page_config(
        page_title="Gestão SSMA Macromaq",
        layout="wide"
    )

    init_db()

    menu = [
        "📋 Nova Inspeção",
        "📊 Indicadores SSMA",
        "✅ Tratativas por Unidade",
        "📂 Histórico de Relatórios"
    ]

    choice = st.sidebar.selectbox("Navegação", menu)

    # =====================================================
    # ABA 1 - NOVA INSPEÇÃO
    # =====================================================
    if choice == "📋 Nova Inspeção":
        col_icon, col_text = st.columns([0.1, 0.9])
        with col_icon:
            # ADAPTAÇÃO GITHUB: Uso de try/except para evitar crash caso a imagem não suba junto
            try:
                st.image(os.path.join(PATH_BASE, "Auditoria - ícon.png"), width=70)
            except:
                st.write("📋")

        with col_text:
            st.title("Auditoria de Visita Técnica SSMA")

        with st.sidebar:
            u_sel = st.selectbox("Unidade", list(UNIDADES.keys()))
            auditor = st.selectbox(
                "Auditor",
                ["Jessica Vieira", "Dilceu Amaral Jr", "Simone Francisco"]
            )
            rel_num = st.text_input(
                "Nº Relatório",
                value=get_next_rel_num(),
                disabled=True
            )
            gerente = st.text_input("Gerente", "Deveide Mafra")

        if 'auditoria' not in st.session_state:
            st.session_state.auditoria = {
                p: {
                    "status": "Conforme",
                    "fotos": [],
                    "obs": "",
                    "prazo": date.today(),
                    "responsavel": "Guilherme Almeida"
                }
                for tema in TEMAS_CHECKLIST.values()
                for p in tema.keys()
            }

        for tema, itens in TEMAS_CHECKLIST.items():
            with st.expander(f"📁 {tema}"):
                for pergunta, nr in itens.items():
                    res = st.radio(
                        f"**{pergunta}**",
                        ["Conforme", "Não Conforme", "Não Aplica"],
                        horizontal=True,
                        key=f"r_{pergunta}"
                    )
                    st.session_state.auditoria[pergunta]["status"] = res

                    if res == "Não Conforme":
                        fotos = st.file_uploader(
                            "Fotos",
                            type=['png', 'jpg', 'jpeg'],
                            accept_multiple_files=True,
                            key=f"up_{pergunta}"
                        )
                        if fotos:
                            st.session_state.auditoria[pergunta]["fotos"] = [
                                f.getvalue() for f in fotos
                            ]

                        st.session_state.auditoria[pergunta]["obs"] = st.text_area(
                            "Diagnóstico:",
                            key=f"obs_{pergunta}"
                        )
                        st.session_state.auditoria[pergunta]["responsavel"] = st.selectbox(
                            "Resp:",
                            ["Guilherme Almeida", "Luiz", "Deveide Mafra"],
                            key=f"resp_{pergunta}"
                        )
                        st.session_state.auditoria[pergunta]["prazo"] = st.date_input(
                            "Prazo:",
                            key=f"praz_{pergunta}"
                        )
                    st.divider()

        if st.button("🚀 FINALIZAR E SALVAR RELATÓRIO", type="primary"):
            try:
                doc = DocxTemplate(os.path.join(PATH_BASE, "JORNADA_SSMA_Template.docx"))
                plano_acao = []
                conformes = []
                constatacoes = []

                for t, its in TEMAS_CHECKLIST.items():
                    for p, nr in its.items():
                        v = st.session_state.auditoria[p]

                        if v["status"] == "Não Conforme":
                            analise = gerar_texto_ia(p, "Não Conforme", v["obs"])
                            why, how = (analise.split("|") + ["Adequar"])[:2]

                            imgs = [
                                InlineImage(
                                    doc,
                                    io.BytesIO(fb),
                                    width=Mm(65)
                                )
                                for fb in v["fotos"]
                            ]

                            constatacoes.append({
                                "setor": p,
                                "nr": nr,
                                "descricao": v["obs"],
                                "fotos": imgs
                            })

                            plano_acao.append({
                                "pergunta": p,
                                "why": why.strip(),
                                "how": how.strip(),
                                "responsavel": v["responsavel"],
                                "prazo": v["prazo"].strftime("%d/%m/%Y")
                            })

                        elif v["status"] == "Conforme":
                            conformes.append({
                                "item": p,
                                "obs": gerar_texto_ia(p, "Conforme")
                            })

                contexto = {
                    "relatorio_num": rel_num,
                    "data_inspecao": datetime.now().strftime("%d/%m/%Y"),
                    "unidade": u_sel,
                    "auditor_selecionado": auditor,
                    "gerente_unidade": gerente,
                    "empreendimento": f"Macromaq - {u_sel}",
                    "local": UNIDADES[u_sel]["ENDERECO"],
                    "constatacoes": constatacoes,
                    "plano_acao": plano_acao,
                    "conformes": conformes
                }

                doc.render(contexto)
                out = io.BytesIO()
                doc.save(out)
                out.seek(0)
                word_binary = out.getvalue()

                conn = sqlite3.connect('jornada_ssma.db')
                c = conn.cursor()
                c.execute(
                    """
                    INSERT INTO inspecoes
                    (rel_num, unidade, data, auditor, arquivo_word)
                    VALUES (?, ?, ?, ?, ?)
                    """,
                    (rel_num, u_sel, datetime.now().strftime("%d/%m/%Y"), auditor, word_binary)
                )
                ins_id = c.lastrowid

                for item_p in plano_acao:
                    c.execute(
                        """
                        INSERT INTO plano_acao
                        (inspecao_id, item, nr, obs, responsavel, prazo, status)
                        VALUES (?, ?, ?, ?, ?, ?, ?)
                        """,
                        (ins_id, item_p["pergunta"], "NR", item_p["why"], item_p["responsavel"], item_p["prazo"], "Pendente")
                    )

                conn.commit()
                conn.close()

                st.success(f"✅ Relatório {rel_num} Gerado e Salvo!")
                st.download_button(
                    "📥 Baixar Agora",
                    word_binary,
                    f"Relatorio_{u_sel}.docx"
                )
            except Exception as e:
                st.error(f"Erro ao gerar documento Word: {e}")

    # =====================================================
    # ABA 2 - DASHBOARD
    # =====================================================
    elif choice == "📊 Indicadores SSMA":
        st.markdown("<h1 style='text-align: center; color: #1c3d5a;'>Dashboard Executivo SSMA</h1>", unsafe_allow_html=True)
        conn = sqlite3.connect('jornada_ssma.db')
        try:
            df = pd.read_sql_query(
                """
                SELECT p.*, i.unidade, i.data as data_auditoria
                FROM plano_acao p
                JOIN inspecoes i
                ON p.inspecao_id = i.id
                """,
                conn
            )
        except:
            df = pd.DataFrame()
        conn.close()

        if not df.empty:
            df['data_auditoria'] = pd.to_datetime(df['data_auditoria'], dayfirst=True)

            with st.sidebar:
                st.header("Filtros")
                u_f = st.selectbox("Unidade:", ["TODAS"] + list(df['unidade'].unique()))
                d_ini = st.date_input("Início:", df['data_auditoria'].min())
                d_fim = st.date_input("Fim:", date.today())

            mask = (df['data_auditoria'].dt.date >= d_ini) & (df['data_auditoria'].dt.date <= d_fim)
            if u_f != "TODAS":
                mask &= (df['unidade'] == u_f)
            df = df[mask]

            total_p = len(df)
            resolvidos = len(df[df['status'] == 'Concluído'])
            eficacia = (resolvidos / total_p * 100) if total_p > 0 else 0

            c1, c2, c3 = st.columns(3)
            with c1:
                fig = go.Figure(go.Indicator(
                    mode="gauge+number",
                    value=eficacia,
                    title={'text': "Eficácia %"},
                    gauge={'bar': {'color': "#2ecc71"}}
                ))
                st.plotly_chart(fig, use_container_width=True)

            with c2:
                st.markdown("<div style='text-align:center;'>", unsafe_allow_html=True)
                try:
                    st.image(os.path.join(PATH_BASE, "Pendências - Icon.png"), width=60)
                except:
                    st.write("⚠️")
                st.markdown(f"""
                    <div style='background-color:white; padding:15px; border-radius:10px; border-left: 8px solid #e74c3c;'>
                        <h4>Pendências</h4>
                        <h1 style='color:#e74c3c;'>{total_p - resolvidos}</h1>
                    </div>
                """, unsafe_allow_html=True)
                st.markdown("</div>", unsafe_allow_html=True)

            with c3:
                st.markdown("<div style='text-align:center;'>", unsafe_allow_html=True)
                try:
                    st.image(os.path.join(PATH_BASE, "Resolvido - Icon.png"), width=60)
                except:
                    st.write("✅")
                st.markdown(f"""
                    <div style='background-color:white; padding:15px; border-radius:10px; border-left: 8px solid #2ecc71;'>
                        <h4>Resolvidos</h4>
                        <h1 style='color:#2ecc71;'>{resolvidos}</h1>
                    </div>
                """, unsafe_allow_html=True)
                st.markdown("</div>", unsafe_allow_html=True)

            st.divider()
            ca, cb = st.columns([1.2, 0.8])

            with ca:
                col_i1, col_t1 = st.columns([0.15, 0.85])
                with col_i1:
                    try:
                        st.image(os.path.join(PATH_BASE, "mapa calor.png"), width=50)
                    except:
                        st.write("📊")
                with col_t1:
                    st.write("### Riscos por Unidade x NR")

                fig_h = px.density_heatmap(
                    df, x="nr", y="unidade", z="id",
                    color_continuous_scale='Reds', text_auto=True
                )
                st.plotly_chart(fig_h, use_container_width=True)

            with cb:
                col_i2, col_t2 = st.columns([0.2, 0.8])
                with col_i2:
                    try:
                        st.image(os.path.join(PATH_BASE, "status açoes.png"), width=50)
                    except:
                        st.write("📈")
                with col_t2:
                    st.write("### Status Ações")

                fig_p = px.pie(
                    df, names='status', hole=.5,
                    color_discrete_sequence=['#e74c3c', '#2ecc71']
                )
                st.plotly_chart(fig_p, use_container_width=True)
        else:
            st.info("Sem dados para o período selecionado.")

    # =====================================================
    # ABA 3 - FOLLOW-UP
    # =====================================================
    elif choice == "✅ Tratativas por Unidade":
        col_f1, col_f2 = st.columns([0.1, 0.9])
        with col_f1:
            try:
                st.image(os.path.join(PATH_BASE, "Follow UP - Icon.png"), width=70)
            except:
                st.write("🔄")
        with col_f2:
            st.title("Follow-up de Pendências")

        conn = sqlite3.connect('jornada_ssma.db')
        try:
            df_p = pd.read_sql_query(
                """
                SELECT p.*, i.unidade
                FROM plano_acao p
                JOIN inspecoes i
                ON p.inspecao_id = i.id
                WHERE p.status = 'Pendente'
                """,
                conn
            )

            if not df_p.empty:
                u_f = st.selectbox("Escolha Unidade:", df_p['unidade'].unique())
                for _, row in df_p[df_p['unidade'] == u_f].iterrows():
                    with st.expander(f"📌 {row['item']}"):
                        st.error(f"Diagnóstico: {row['obs']}")
                        f_dep = st.file_uploader(f"Anexar Foto #{row['id']}", key=f"f_{row['id']}")

                        if st.button(f"Validar #{row['id']}", type="primary"):
                            if f_dep:
                                cursor = conn.cursor()
                                cursor.execute(
                                    """
                                    UPDATE plano_acao
                                    SET status='Concluído', foto_depois=?
                                    WHERE id=?
                                    """,
                                    (f_dep.getvalue(), row['id'])
                                )
                                conn.commit()
                                st.rerun()
            else:
                st.success("Tudo em conformidade!")
        except:
            st.info("Aguardando primeiras inspeções.")
        conn.close()

    # =====================================================
    # ABA 4 - HISTÓRICO
    # =====================================================
    elif choice == "📂 Histórico de Relatórios":
        col_h1, col_h2 = st.columns([0.1, 0.9])
        with col_h1:
            try:
                st.image(os.path.join(PATH_BASE, "Auditoria - ícon.png"), width=70)
            except:
                st.write("📂")
        with col_h2:
            st.title("Histórico de Relatórios")

        conn = sqlite3.connect('jornada_ssma.db')
        try:
            df_hist = pd.read_sql_query(
                """
                SELECT id, rel_num, unidade, data, auditor
                FROM inspecoes
                ORDER BY id DESC
                """,
                conn
            )

            for _, row in df_hist.iterrows():
                c1, c2, c3, c4 = st.columns([2, 2, 2, 2])
                c1.write(f"**Nº:** {row['rel_num']}")
                c2.write(f"**Unidade:** {row['unidade']}")
                c3.write(f"**Data:** {row['data']}")

                cursor = conn.cursor()
                cursor.execute("SELECT arquivo_word FROM inspecoes WHERE id=?", (row['id'],))
                blob = cursor.fetchone()[0]

                c4.download_button(
                    "📥 Baixar",
                    blob,
                    f"Relatorio_{row['rel_num'].replace('/','-')}.docx",
                    key=f"h_{row['id']}"
                )
                st.divider()
        except:
            st.info("Histórico vazio.")
        conn.close()

if __name__ == "__main__":
    main()
