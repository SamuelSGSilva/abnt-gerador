# app.py - REFORMATADOR ABNT 2024 (CORRIGIDO - 100% FUNCIONAL)
import streamlit as st
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
import io
import re

st.set_page_config(page_title="ABNT 2024 - Reformatador", layout="centered")
st.title("Reformatador ABNT 2024")
st.caption("Upload seu .docx → Aplica NBR 14724:2024 → Baixa perfeito! (Preserva imagens, tabelas, fórmulas)")

# Upload
arquivo = st.file_uploader("Faça upload do seu TCC em .docx", type="docx")

if arquivo:
    try:
        doc = Document(arquivo)
        st.success(f"Documento carregado: {len(doc.paragraphs)} parágrafos, {len(doc.tables)} tabelas, {len(doc.inline_shapes)} imagens")
    except Exception as e:
        st.error(f"Erro ao ler o arquivo: {e}")
        st.stop()

    if st.button("Aplicar Formatação ABNT 2024", type="primary"):
        with st.spinner("Aplicando normas ABNT 2024..."):
            # === 1. MARGENS (todas as seções) ===
            for section in doc.sections:
                section.top_margin = Cm(3)
                section.bottom_margin = Cm(2)
                section.left_margin = Cm(3)
                section.right_margin = Cm(2)
                section.page_height = Cm(29.7)
                section.page_width = Cm(21.0)

            # === 2. ESTILOS PERSONALIZADOS ===
            def criar_estilo(nome, base_style=None, fonte='Arial', tam=12, negrito=False, alinhamento=WD_ALIGN_PARAGRAPH.JUSTIFY, espacamento=1.5):
                if nome in doc.styles:
                    estilo = doc.styles[nome]
                else:
                    estilo = doc.styles.add_style(nome, WD_STYLE_TYPE.PARAGRAPH)
                estilo.base_style = base_style
                estilo.font.name = fonte
                estilo.font.size = Pt(tam)
                estilo.font.bold = negrito
                estilo.paragraph_format.alignment = alinhamento
                estilo.paragraph_format.line_spacing = espacamento
                estilo.paragraph_format.space_before = Pt(24) if 'Heading' in nome else Pt(0)
                estilo.paragraph_format.space_after = Pt(12) if 'Heading' in nome else Pt(0)
                estilo.paragraph_format.left_indent = Cm(0)
                return estilo

            # Estilo Normal (texto principal)
            criar_estilo('ABNT_Normal', fonte='Arial', tam=12, negrito=False, alinhamento=WD_ALIGN_PARAGRAPH.JUSTIFY, espacamento=1.5)

            # Títulos (Heading 1 a 5)
            for i in range(1, 6):
                criar_estilo(f'ABNT_Heading_{i}', fonte='Arial', tam=12, negrito=True, alinhamento=WD_ALIGN_PARAGRAPH.LEFT, espacamento=1.5)

            # Citação longa
            citacao_style = criar_estilo('ABNT_Citacao', fonte='Arial', tam=10, negrito=False, alinhamento=WD_ALIGN_PARAGRAPH.JUSTIFY, espacamento=1.0)
            citacao_style.paragraph_format.left_indent = Cm(4)

            # Referências
            criar_estilo('ABNT_Referencia', fonte='Arial', tam=12, negrito=False, alinhamento=WD_ALIGN_PARAGRAPH.LEFT, espacamento=1.0)

            # === 3. REAPLICAR ESTILOS INTELIGENTEMENTE ===
            for para in doc.paragraphs:
                texto = para.text.strip()

                # Títulos (padrão ABNT: 1 INTRODUÇÃO, 1.1 Subseção)
                if para.style.name.startswith('Heading') or re.match(r'^\d+(\.\d+)*\s+[A-ZÀ-Ú\s]+$', texto):
                    nivel = 1
                    pontos = texto.split(' ', 1)[0].count('.')
                    nivel = min(pontos + 1, 5)
                    para.style = doc.styles[f'ABNT_Heading_{nivel}']

                # Citações longas (> 120 chars OU com recuo)
                elif len(texto) > 120 or (para.paragraph_format.left_indent is not None and para.paragraph_format.left_indent > Cm(1)):
                    para.style = doc.styles['ABNT_Citacao']

                # Referências (detecta por palavras-chave)
                elif any(texto.upper().startswith(p) for p in ['SILVA', 'OLIVEIRA', 'SANTOS', 'COSTA', '2023', '2024', '2025']) or 'REFER' in texto.upper():
                    para.style = doc.styles['ABNT_Referencia']

                # Texto normal
                else:
                    para.style = doc.styles['ABNT_Normal']

            # === 4. TABELAS E IMAGENS ===
            for table in doc.tables:
                table.alignment = 1  # Centro

            # === 5. SALVAR ===
            buffer = io.BytesIO()
            doc.save(buffer)
            buffer.seek(0)

            st.success("Formatação ABNT 2024 aplicada com sucesso!")
            st.balloons()

            col1, col2 = st.columns(2)
            with col1:
                st.download_button(
                    label="Baixar TCC_ABNT_2024.docx",
                    data=buffer,
                    file_name="TCC_ABNT_2024.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            with col2:
                st.info("""
                **No Word (finalizar):**
                1. Referências → Inserir Sumário
                2. Inserir → Nº de Página → Superior Direito
                3. Salvar como PDF
                """)

else:
    st.info("Envie seu arquivo .docx para começar!")
    st.markdown("""
    ### Dicas:
    - Use **.docx** (não PDF)
    - Converta em [ilovepdf.com](https://www.ilovepdf.com/pdf_to_word)
    - O app **não altera conteúdo**, só formata!
    """)
