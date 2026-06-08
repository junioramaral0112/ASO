import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
from datetime import datetime
import io

# --- CONFIGURAÇÕES E MAPEAMENTO ---
CHECKLIST_NR = {
    "Equipamentos de Proteção Individual (EPIs)": "NR 06",
    "Máquinas e Ferramentas (Proteção)": "NR 12",
    "Máquinas e Ferramentas (Manutenção)": "NR 12",
    "Organização e Ambiente (5S)": "NR 11 / NR 17",
    "Iluminação e Ventilação": "NR 17",
    "Prevenção de Incêndios (Extintores)": "NR 23",
    "Rotas de Fuga e Emergência": "NR 23",
    "Armazenamento de Produtos Químicos": "NR 20 / NR 26",
    "Fichas de Segurança (FDS)": "NR 26",
    "Vasos de Pressão (Compressores)": "NR 13"
}

def main():
    st.set_page_config(page_title="Sistema de Auditoria SSMA", layout="wide")
    st.title("🛡️ Sistema de Auditoria - Jornada SSMA")

    # --- SESSÃO DE DADOS ---
    if 'auditoria_data' not in st.session_state:
        st.session_state.auditoria_data = {}

    # --- SIDEBAR: CONFIGURAÇÕES ---
    with st.sidebar:
        st.header("Dados da Unidade")
        unidade = st.selectbox("Unidade/Filial", ["Joinville", "Curitiba", "São José", "Itupeva"])
        auditor = st.selectbox("Auditor Responsável", ["Luciano Fogaça", "Jessica Vieira", "Dilceu Amaral Jr"])
        gerente = st.text_input("Gerente da Unidade", "Sr. Deveide Mafra")
        relatorio_num = st.text_input("Relatório Nº", "001/2026")
        data_visita = st.date_input("Data da Visita", datetime.now())

    # --- CHECKLIST ---
    st.header(f"Checklist de Auditoria - {unidade}")
    
    for item, nr in CHECKLIST_NR.items():
        with st.expander(f"{item} ({nr})"):
            col1, col2 = st.columns([1, 2])
            
            with col1:
                status = st.radio("Status", ["Conforme", "Não Conforme", "Não Aplica"], key=f"status_{item}")
            
            if status == "Não Conforme":
                with col2:
                    # Lógica de múltiplas fotos
                    st.write("**Evidências Fotográficas**")
                    if item not in st.session_state.auditoria_data:
                        st.session_state.auditoria_data[item] = {"fotos": [], "descricao": "", "prazo": datetime.now()}
                    
                    foto = st.camera_input(f"Tirar foto para {item}", key=f"cam_{item}")
                    if foto:
                        # Armazena os bytes da foto
                        st.session_state.auditoria_data[item]["fotos"].append(foto.getvalue())
                        st.success(f"Foto adicionada! Total: {len(st.session_state.auditoria_data[item]['fotos'])}")

                    # Visualização das fotos atuais
                    if st.session_state.auditoria_data[item]["fotos"]:
                        cols_img = st.columns(4)
                        for i, img_bytes in enumerate(st.session_state.auditoria_data[item]["fotos"]):
                            cols_img[i % 4].image(img_bytes, width=100)
                            if cols_img[i % 4].button("Remover", key=f"del_{item}_{i}"):
                                st.session_state.auditoria_data[item]["fotos"].pop(i)
                                st.rerun()

                    # Dados do Plano de Ação
                    st.session_state.auditoria_data[item]["descricao"] = st.text_area("Descrição do Risco/Ação", key=f"desc_{item}")
                    st.session_state.auditoria_data[item]["responsavel"] = st.text_input("Responsável pela Correção", key=f"resp_{item}")
                    st.session_state.auditoria_data[item]["prazo"] = st.date_input("Prazo", key=f"date_{item}")

    # --- GERAÇÃO DO RELATÓRIO ---
    st.divider()
    if st.button("🚀 GERAR RELATÓRIO FINAL"):
        try:
            doc = DocxTemplate("JORNADA_SSMA_Template.docx") # Seu arquivo com as tags
            
            constatacoes = []
            plano_acao = []

            for item, dados in st.session_state.auditoria_data.items():
                if len(dados["fotos"]) > 0 or dados["descricao"] != "":
                    # Prepara as fotos para o InlineImage do docxtpl
                    fotos_inline = [InlineImage(doc, io.BytesIO(f), width=Mm(75)) for f in dados["fotos"]]
                    
                    # Dados para a seção de constatações
                    constatacoes.append({
                        "setor": item,
                        "nr": CHECKLIST_NR[item],
                        "descricao": dados["descricao"],
                        "fotos": fotos_inline
                    })
                    
                    # Dados para a tabela de plano de ação
                    plano_acao.append({
                        "descricao": dados["descricao"],
                        "acao": f"Adequar conforme {CHECKLIST_NR[item]}",
                        "responsavel": dados["responsavel"],
                        "prazo": dados["prazo"].strftime("%d/%m/%Y")
                    })

            contexto = {
                "relatorio_num": relatorio_num,
                "data_inspecao": data_visita.strftime("%d/%m/%Y"),
                "unidade": unidade,
                "auditor_selecionado": auditor,
                "gerente_unidade": gerente,
                "constatacoes": constatacoes,
                "plano_acao": plano_acao
            }

            doc.render(contexto)
            
            # Saída do arquivo
            target_stream = io.BytesIO()
            doc.save(target_stream)
            target_stream.seek(0)

            st.download_button(
                label="📥 Baixar Relatório Gerado",
                data=target_stream,
                file_name=f"Relatorio_SSMA_{unidade}_{relatorio_num.replace('/','-')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            st.success("Relatório gerado com sucesso!")

        except Exception as e:
            st.error(f"Erro ao gerar relatório: {e}")

if __name__ == "__main__":
    main()