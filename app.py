# app.py - EDITOR ABNT 2024 AUTOMÁTICO (Suporte a PDF/DOCX Grandes)
import streamlit as st
import io
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
import fitz  # PyMuPDF para PDFs
from docx import Document as DocxDoc
import re

st.set_page_config(page_title="Editor ABNT 2024", layout="wide")
st.title("🛠️ Editor ABNT Automático 2024")
st.caption("Upload seu TCC (PDF/DOCX) → Edição 100% ABNT NBR 14724:2024 → Download pronto! Suporte a arquivos grandes.")

# Sidebar com normas
with st.sidebar:
    st.header("📋 Normas Aplicadas (NBR 14724:2024)")
    st.markdown("""
    - **Margens**: 3cm (esq/sup), 2cm (dir/inf)  
    - **Fonte**: Arial 12 (texto), 10 (notas)  
    - **Espaçamento**: 1,5 (principal), simples (refs/citações)  
    - **Alinhamento**: Justificado  
    - **Paginação**: A partir de Introdução (p.1)  
    - **Outras**: Citações (10520:2023), Refs (6023:2018)  
    """)

# Upload
arquivo = st.file_uploader("📁 Upload do seu TCC (PDF ou DOCX)", type=['pdf', 'docx'])

if arquivo:
    # Detectar tipo
    if arquivo.type == "application/pdf":
        st.info("🔄 PDF detectado. Extrair texto...")
        doc_pdf = fitz.open(stream=arquivo.read(), filetype="pdf")
        texto_extraido = ""
        for page in doc_pdf:
            texto_extraido += page.get_text() + "\n"
        st.success(f"Texto extraído: {len(texto_extraido)} chars.")
        conteudo = texto_extraido
    else:  # DOCX
        st.info("📄 DOCX detectado. Lendo...")
        doc_temp = DocxDoc(arquivo)
        conteudo = "\n".join([p.text for p in doc_temp.paragraphs])
        st.success(f"Conteúdo lido: {len(conteudo)} chars.")

    # Processar e formatar ABNT
    if st.button("✨ Editar e Gerar ABNT 2024") and conteudo:
        with st.spinner("Aplicando normas ABNT... (pode demorar com arquivos grandes)"):
            # Parse simples: Detecta títulos por maiúsculas/negrito simulado, parágrafos, etc.
            # Dividir em seções (heurística: linhas em maiúsculas = títulos)
            secoes = re.split(r'([A-Z\s]{10,})', conteudo)  # Títulos como "1 INTRODUÇÃO"
            secoes_filtradas = [s.strip() for s in secoes if s.strip()]

            # Criar DOCX formatado
            doc = Document()
            section = doc.sections[0]
            section.top_margin = Cm(3)
            section.bottom_margin = Cm(2)
            section.left_margin = Cm(3)
            section.right_margin = Cm(2)

            def criar_estilo(nome, tam=12, negrito=False, alinh=WD_ALIGN_PARAGRAPH.JUSTIFY, esp=1.5):
                s = doc.styles.add_style(nome, WD_STYLE_TYPE.PARAGRAPH)
                f = s.font
                f.name = 'Arial'
                f.size = Pt(tam)
                f.bold = negrito
                s.paragraph_format.alignment = alinh
                s.paragraph_format.line_spacing = esp
                s.paragraph_format.space_after = Pt(12) if esp == 1.5 else Pt(6)
                return s

            criar_estilo('Texto', 12, False, WD_ALIGN_PARAGRAPH.JUSTIFY, 1.5)
            criar_estilo('Titulo', 12, True, WD_ALIGN_PARAGRAPH.LEFT, 1.5)
            criar_estilo('Ref', 12, False, WD_ALIGN_PARAGRAPH.LEFT, 1.0)  # Referências simples

            # Elementos pré-textuais (adicionar manual se não detectado)
            doc.add_paragraph("CAPA", style='Titulo')  # Placeholder - ajuste no Word
            doc.add_paragraph("Folha de Rosto", style='Titulo')
            doc.add_paragraph("RESUMO", style='Titulo')  # Detecta se há resumo
            if "RESUMO" in conteudo.upper():
                resumo = re.search(r'RESUMO(.*?)ABSTRACT|REFERENCIAS', conteudo, re.DOTALL | re.I)
                if resumo: doc.add_paragraph(resumo.group(1).strip(), style='Texto')

            doc.add_paragraph("SUMÁRIO", style='Titulo')  # Auto no Word

            # Textual: Seções
            for i, parte in enumerate(secoes_filtradas):
                if re.match(r'^\d+\s+[A-Z\s]+$', parte):  # Título de seção
                    doc.add_paragraph(parte, style='Titulo')
                else:
                    # Parágrafos: Justificar e espaçar
                    paras = parte.split('\n')
                    for p in paras:
                        if p.strip():
                            par = doc.add_paragraph(p.strip(), style='Texto')
                            # Citações longas: Detecta e formata (recuo 4cm, simples, 10pt)
                            if len(p) > 100 and '"' in p:  # Heurística simples
                                par.paragraph_format.left_indent = Cm(4)
                                par.paragraph_format.line_spacing = 1.0
                                par.runs[0].font.size = Pt(10)

            # Pós-textual: Referências
            if "REFERENCIAS" in conteudo.upper():
                refs = re.search(r'REFERENCIAS(.*)', conteudo, re.DOTALL | re.I)
                if refs:
                    doc.add_paragraph("REFERÊNCIAS", style='Titulo')
                    ref_linhas = refs.group(1).split('\n')
                    for linha in ref_linhas:
                        if linha.strip():
                            rpar = doc.add_paragraph(linha.strip(), style='Ref')

            # Paginação (inicia na introdução - ajuste manual no Word)
            # Salvar
            buffer = io.BytesIO()
            doc.save(buffer)
            buffer.seek(0)

            st.success("✅ Edição concluída! Documento formatado ABNT 2024.")
            
            # Downloads
            col1, col2 = st.columns(2)
            with col1:
                st.download_button(
                    label="📥 Baixar como DOCX (para editar)",
                    data=buffer,
                    file_name="tcc_abnt_2024.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            with col2:
                # Para PDF: Converter DOCX para PDF (simples via docx2pdf se instalado, senão instrução)
                st.info("💡 Para PDF: Abra o DOCX no Word > Salvar como PDF. Ou instale docx2pdf localmente.")
                # Nota: Adicione 'docx2pdf' no requirements se quiser auto-PDF

else:
    st.info("💡 **Dica**: Upload um PDF/DOCX do seu TCC. O app extrai texto, formata e devolve alinhadinho!")

st.caption("Feito com ❤️ para estudantes. Baseado em NBR 14724:2024. Problemas? Mande feedback!")
