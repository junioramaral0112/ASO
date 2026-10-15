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
# CONFIGURAÇÃO GEMINI IA (STREAMLIT SECRETS / AMBIENTE)
# =========================================================
if "GEMINI_API_KEY" in st.secrets:
    genai.configure(api_key=st.secrets["GEMINI_API_KEY"])
else:
    genai.configure(api_key=os.environ.get("GEMINI_API_KEY", ""))

model = genai.GenerativeModel("gemini-1.5-flash")

PATH_BASE = ""

def gerar_texto_ia(pergunta, status, diagnostico=""):
    try:
        if status == "Conforme":
            prompt = f"Gere uma frase curta técnica confirmando que o item '{pergunta}' cumpre os requisitos SSMA da Macromaq."
        else:
            prompt = f"Especialista SSMA. Item: {pergunta}. Diagnóstico: {diagnostico}. Formato: JUSTIFICATIVA | AÇÃO CORRETIVA."
        response = model.generate_content(prompt)
        return response.text.strip()
    except Exception:
        return "Item em conformidade." if status == "Conforme" else "Risco identificado | Realizar adequação conforme NR."

# =========================================================
# BANCO DE DADOS
# =========================================================
def init_db():
    conn = sqlite3.connect("jornada_ssma.db")
    c = conn.cursor()

    c.execute("""
        CREATE TABLE IF NOT EXISTS inspecoes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            rel_num TEXT,
            unidade TEXT,
            data TEXT,
            auditor TEXT,
            arquivo_word BLOB
        )
    """)

    c.execute("""
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
    """)

    conn.commit()
    conn.close()

def get_next_rel_num():
    ano = datetime.now().year
    try:
        conn = sqlite3.connect("jornada_ssma.db")
        c = conn.cursor()
        c.execute(
            "SELECT COUNT(*) FROM inspecoes WHERE rel_num LIKE ?",
            (f"%/{ano}",)
        )
        count = c.fetchone()[0]
        conn.close()
        return f"{(count + 1):04d}/{ano}"
    except Exception:
        return f"0001/{ano}"

# =========================================================
# CONFIGURAÇÕES DE NEGÓCIO - FILIAIS E CHECKLIST
# =========================================================
UNIDADES = {
    "SÃO JOSÉ (MATRIZ)": {
        "CNPJ": "83.675.413/0001-01",
        "ENDERECO": "BR 101, km 210 / Bairro: Picadas do Sul – São José – SC"
    },
    "CHAPECÓ": {
        "CNPJ": "83.675.413/0002-84",
        "ENDERECO": "Rua Xanxerê, 360E – Bairro Líder – Chapecó – SC"
    },
    "CRICIÚMA": {
        "CNPJ": "83.675.413/0003-65",
        "ENDERECO": "Rodovia Jorge Lacerda, 4555 – Bairro Sangão – Criciúma – SC"
    },
    "LAGES": {
        "CNPJ": "83.675.413/0004-46",
        "ENDERECO": "Av. Dom Pedro II, 2887 – São Cristóvão – Lages – SC"
    },
    "CURITIBA": {
        "CNPJ": "83.675.413/0005-27",
        "ENDERECO": "BR 116 (Linha Verde), 17350 – Pinheirinho – Curitiba – PR"
    },
    "CASCAVEL": {
        "CNPJ": "83.675.413/0006-08",
        "ENDERECO": "BR 277, Km 594 – Trevo Cataratas – Cascavel – PR"
    },
    "LONDRINA": {
        "CNPJ": "83.675.413/0007-90",
        "ENDERECO": "Av. Tiradentes, 5500 – Jardim Rosicler – Londrina – PR"
    },
    "PORTO ALEGRE": {
        "CNPJ": "83.675.413/0008-71",
        "ENDERECO": "Av. Severo Dullius, 1395 – São João – Porto Alegre – RS"
    },
    "PASSO FUNDO": {
        "CNPJ": "83.675.413/0009-52",
        "ENDERECO": "RS 153, Km 02 – Nossa Senhora Aparecida – Passo Fundo – RS"
    },
    "CAXIAS DO SUL": {
        "CNPJ": "83.675.413/0010-94",
        "ENDERECO": "BR 116, Km 144 – São Ciro – Caxias do Sul – RS"
    },
    "JOINVILLE": {
        "CNPJ": "83.675.413/0011-75",
        "ENDERECO": "BR 101, KM 17 – Pirabeiraba – Joinville – SC"
    },
    "MARINGÁ": {
        "CNPJ": "83.675.413/0012-56",
        "ENDERECO": "PR-317, Km 05 – Parque Industrial – Maringá – PR"
    },
    "PELOTAS": {
        "CNPJ": "83.675.413/0013-37",
        "ENDERECO": "BR 116, Km 524 – Três Vendas – Pelotas – RS"
    },
    "SANTA MARIA": {
        "CNPJ": "83.675.413/0014-18",
        "ENDERECO": "BR 287, Km 242 – Camobi – Santa Maria – RS"
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
        "Higiene: Lava-olhos higienizado e funcional?": "NR 24",
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

    # -----------------------------------------------------
    # ABA 1 - NOVA INSPEÇÃO
    # -----------------------------------------------------
    if choice == "📋 Nova Inspeção":
        col_icon, col_text = st.columns([0.1, 0.9])
        with col_icon:
            try:
                st.image(os.path.join(PATH_BASE, "Auditoria - ícon.png"), width=70)
            except Exception:
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
            gerente = st.text_input("Gerente da Unidade", "Deveide Mafra")

        if "auditoria" not in st.session_state:
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
                            "Fotos da Evidência",
                            type=["png", "jpg", "jpeg"],
                            accept_multiple_files=True,
                            key=f"up_{pergunta}"
                        )
                        if fotos:
                            st.session_state.auditoria[pergunta]["fotos"] = [
                                f.getvalue() for f in fotos
                            ]

                        st.session_state.auditoria[pergunta]["obs"] = st.text_area(
                            "Diagnóstico / Apontamento:",
                            key=f"obs_{pergunta}"
                        )
                        st.session_state.auditoria[pergunta]["responsavel"] = st.selectbox(
                            "Responsável:",
                            ["Guilherme Almeida", "Luiz", "Deveide Mafra"],
                            key=f"resp_{pergunta}"
                        )
                        st.session_state.auditoria[pergunta]["prazo"] = st.date_input(
                            "Prazo Limite:",
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
                            partes = analise.split("|")
                            why = partes[0] if len(partes) > 0 else "Risco identificado"
                            how = partes[1] if len(partes) > 1 else "Adequar conforme NR"

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
                    "cnpj": UNIDADES[u_sel]["CNPJ"],
                    "constatacoes": constatacoes,
                    "plano_acao": plano_acao,
                    "conformes": conformes
                }

                doc.render(contexto)
                out = io.BytesIO()
                doc.save(out)
                out.seek(0)
                word_binary = out.getvalue()

                conn = sqlite3.connect("jornada_ssma.db")
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

                st.success(f"✅ Relatório {rel_num} gerado e salvo com sucesso!")
                st.download_button(
                    "📥 Baixar Relatório Word",
                    word_binary,
                    f"Relatorio_{u_sel}_{rel_num.replace('/', '-')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            except Exception as e:
                st.error(f"Erro ao processar e salvar o relatório: {e}")

    # -----------------------------------------------------
    # ABA 2 - DASHBOARD
    # -----------------------------------------------------
    elif choice == "📊 Indicadores SSMA":
        st.markdown("<h1 style='text-align: center; color: #1c3d5a;'>Dashboard Executivo SSMA</h1>", unsafe_allow_html=True)
        conn = sqlite3.connect("jornada_ssma.db")
        try:
            df = pd.read_sql_query(
                """
                SELECT p.*, i.unidade, i.data as data_auditoria
                FROM plano_acao p
                JOIN inspecoes i ON p.inspecao_id = i.id
                """,
                conn
            )
        except Exception:
            df = pd.DataFrame()
        conn.close()

        if not df.empty:
            df["data_auditoria"] = pd.to_datetime(df["data_auditoria"], dayfirst=True)

            with st.sidebar:
                st.header("Filtros")
                u_f = st.selectbox("Unidade:", ["TODAS"] + list(UNIDADES.keys()))
                d_ini = st.date_input("Data Início:", df["data_auditoria"].min())
                d_fim = st.date_input("Data Fim:", date.today())

            mask = (df["data_auditoria"].dt.date >= d_ini) & (df["data_auditoria"].dt.date <= d_fim)
            if u_f != "TODAS":
                mask &= (df["unidade"] == u_f)
            df = df[mask]

            total_p = len(df)
            resolvidos = len(df[df["status"] == "Concluído"])
            eficacia = (resolvidos / total_p * 100) if total_p > 0 else 0

            c1, c2, c3 = st.columns(3)
            with c1:
                fig = go.Figure(go.Indicator(
                    mode="gauge+number",
                    value=eficacia,
                    title={"text": "Eficácia (%)"},
                    gauge={"bar": {"color": "#2ecc71"}}
                ))
                st.plotly_chart(fig, use_container_width=True)

            with c2:
                st.markdown(f"""
                    <div style='background-color:white; padding:20px; border-radius:10px; border-left: 8px solid #e74c3c; text-align:center;'>
                        <h4 style='margin:0; color:#555;'>Pendências Abertas</h4>
                        <h1 style='margin:0; color:#e74c3c; font-size:48px;'>{total_p - resolvidos}</h1>
                    </div>
                """, unsafe_allow_html=True)

            with c3:
                st.markdown(f"""
                    <div style='background-color:white; padding:20px; border-radius:10px; border-left: 8px solid #2ecc71; text-align:center;'>
                        <h4 style='margin:0; color:#555;'>Ações Concluídas</h4>
                        <h1 style='margin:0; color:#2ecc71; font-size:48px;'>{resolvidos}</h1>
                    </div>
                """, unsafe_allow_html=True)

            st.divider()
            ca, cb = st.columns([1.2, 0.8])

            with ca:
                st.write("### Ocorrências por Unidade")
                fig_bar = px.histogram(df, x="unidade", color="status", barmode="group",
                                       color_discrete_map={"Pendente": "#e74c3c", "Concluído": "#2ecc71"})
                st.plotly_chart(fig_bar, use_container_width=True)

            with cb:
                st.write("### Distribuição dos Status")
                fig_p = px.pie(
                    df, names="status", hole=0.5,
                    color="status",
                    color_discrete_map={"Pendente": "#e74c3c", "Concluído": "#2ecc71"}
                )
                st.plotly_chart(fig_p, use_container_width=True)
        else:
            st.info("Nenhuma inspeção cadastrada no momento.")

    # -----------------------------------------------------
    # ABA 3 - FOLLOW-UP
    # -----------------------------------------------------
    elif choice == "✅ Tratativas por Unidade":
        col_f1, col_f2 = st.columns([0.1, 0.9])
        with col_f1:
            try:
                st.image(os.path.join(PATH_BASE, "Follow UP - Icon.png"), width=70)
            except Exception:
                st.write("🔄")
        with col_f2:
            st.title("Follow-up de Pendências")

        conn = sqlite3.connect("jornada_ssma.db")
        try:
            df_p = pd.read_sql_query(
                """
                SELECT p.*, i.unidade
                FROM plano_acao p
                JOIN inspecoes i ON p.inspecao_id = i.id
                WHERE p.status = 'Pendente'
                """,
                conn
            )

            if not df_p.empty:
                unidades_disponiveis = df_p["unidade"].unique()
                u_f = st.selectbox("Selecione a Unidade:", unidades_disponiveis)
                itens_filtrados = df_p[df_p["unidade"] == u_f]

                for _, row in itens_filtrados.iterrows():
                    with st.expander(f"📌 {row['item']} | Prazo: {row['prazo']}"):
                        st.error(f"**Diagnóstico:** {row['obs']}")
                        st.write(f"**Responsável:** {row['responsavel']}")
                        f_dep = st.file_uploader(f"Anexar Comprovante / Foto de Baixa #{row['id']}", key=f"f_{row['id']}")

                        if st.button(f"Baixar Pendência #{row['id']}", type="primary", key=f"btn_{row['id']}"):
                            cursor = conn.cursor()
                            foto_bytes = f_dep.getvalue() if f_dep else None
                            cursor.execute(
                                """
                                UPDATE plano_acao
                                SET status='Concluído', foto_depois=?
                                WHERE id=?
                                """,
                                (foto_bytes, row["id"])
                            )
                            conn.commit()
                            st.success(f"Item #{row['id']} concluído com sucesso!")
                            st.rerun()
            else:
                st.success("🎉 Todas as filiais estão 100% em conformidade! Nenhuma pendência em aberto.")
        except Exception as e:
            st.error(f"Erro ao carregar pendências: {e}")
        finally:
            conn.close()

    # -----------------------------------------------------
    # ABA 4 - HISTÓRICO
    # -----------------------------------------------------
    elif choice == "📂 Histórico de Relatórios":
        col_h1, col_h2 = st.columns([0.1, 0.9])
        with col_h1:
            try:
                st.image(os.path.join(PATH_BASE, "Auditoria - ícon.png"), width=70)
            except Exception:
                st.write("📂")
        with col_h2:
            st.title("Histórico de Relatórios")

        conn = sqlite3.connect("jornada_ssma.db")
        try:
            df_hist = pd.read_sql_query(
                """
                SELECT id, rel_num, unidade, data, auditor
                FROM inspecoes
                ORDER BY id DESC
                """,
                conn
            )

            if not df_hist.empty:
                for _, row in df_hist.iterrows():
                    c1, c2, c3, c4 = st.columns([2, 2, 2, 2])
                    c1.write(f"**Nº:** {row['rel_num']}")
                    c2.write(f"**Unidade:** {row['unidade']}")
                    c3.write(f"**Data:** {row['data']}")

                    cursor = conn.cursor()
                    cursor.execute("SELECT arquivo_word FROM inspecoes WHERE id=?", (row["id"],))
                    dado = cursor.fetchone()
                    blob = dado[0] if dado else None

                    if blob:
                        c4.download_button(
                            "📥 Baixar DOCX",
                            blob,
                            f"Relatorio_{row['rel_num'].replace('/', '-')}.docx",
                            key=f"h_{row['id']}"
                        )
                    st.divider()
            else:
                st.info("Nenhum relatório emitido até o momento.")
        except Exception:
            st.info("Histórico ainda vazio.")
        finally:
            conn.close()

if __name__ == "__main__":
    main()
