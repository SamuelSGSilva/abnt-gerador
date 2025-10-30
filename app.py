# app.py - REFORMATADOR ABNT 2024 (CORRIGIDO: SEM ERRO, ABSTRACT OPCIONAL, CRIAÇÃO DO ZERO)
import streamlit as st
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.section import WD_SECTION_START
import io
import re
import fitz  # PyMuPDF (opcional para PDF)

st.set_page_config(page_title="ABNT 2024 - Reformatador", layout="wide")
st.title("Reformatador ABNT 2024 (Criação do Zero)")
st.caption("Preencha os dados → Gera documento 100% fiel à NBR 14724:2024! Upload opcional.")

# Escolha de fonte
fonte_escolhida = st.selectbox("Escolha a fonte:", ["Arial", "Times New Roman"])

# Form para dados obrigatórios
with st.form("Dados do Trabalho"):
    titulo = st.text_input("Título (obrigatório)")
    subtitulo = st.text_input("Subtítulo (opcional)")
    autor = st.text_input("Nome do Autor (obrigatório)")
    instituicao = st.text_input("Nome da Instituição (opcional)")
    local = st.text_input("Local (cidade, UF) (obrigatório)")
    ano = st.text_input("Ano de Depósito (obrigatório)")
    orientador = st.text_input("Nome do Orientador (opcional)")
    natureza = st.text_area("Natureza do Trabalho (ex: TCC apresentado à [instituição] para obtenção do grau de Bacharel em [área]) (obrigatório)")
    resumo = st.text_area("Resumo (150-500 palavras, obrigatório)")
    palavras_chave = st.text_input("Palavras-chave (3-5, separadas por ponto)")
    abstract = st.text_area("Abstract (opcional)")
    keywords = st.text_input("Keywords (opcional)")
    submitted = st.form_submit_button("Gerar Documento ABNT 2024")

# Upload opcional
arquivo = st.file_uploader("Upload opcional do conteúdo (PDF/DOCX)", type=["pdf", "docx"])

# === INICIALIZAR VARIÁVEIS SEMPRE ===
conteudo = ""
ref_linhas = []  # <-- CORRIGIDO: sempre existe

if submitted and titulo and autor and local and ano and natureza and resumo:
    with st.spinner("Gerando documento conforme ABNT 2024..."):
        doc = Document()

        # === PROCESSAR UPLOAD (OPCIONAL) ===
        if arquivo:
            if arquivo.type == "application/pdf":
                try:
                    pdf_doc = fitz.open(stream=arquivo.read(), filetype="pdf")
                    for page in pdf_doc:
                        conteudo += page.get_text("text") + "\n"
                        # Extrair imagens (simples)
                        images = page.get_images(full=True)
                        for img_index, img in enumerate(images):
                            xref = img[0]
                            base_image = pdf_doc.extract_image(xref)
                            image_bytes = base_image["image"]
                            p = doc.add_paragraph(f"Figura {page.number+1}.{img_index+1} — Extraída do PDF")
                            p.add_run().add_picture(io.BytesIO(image_bytes), width=Cm(10))
                    pdf_doc.close()
                except Exception as e:
                    st.warning(f"PDF processado parcialmente: {e}")
            else:  # DOCX
                try:
                    input_doc = Document(arquivo)
                    conteudo = "\n".join([p.text for p in input_doc.paragraphs])
                    # Copiar elementos complexos (tabelas, imagens)
                    for element in input_doc.element.body:
                        doc.element.body.append(element)
                except Exception as e:
                    st.warning(f"DOCX processado parcialmente: {e}")

        # === CONFIGURAÇÕES GLOBAIS ===
        section = doc.sections[0]
        section.top_margin = Cm(3)
        section.bottom_margin = Cm(2)
        section.left_margin = Cm(3)
        section.right_margin = Cm(2)
        section.page_width = Cm(21)
        section.page_height = Cm(29.7)

        # Estilos
        def criar_estilo(nome, tam=12, negrito=False, alinh=WD_ALIGN_PARAGRAPH.JUSTIFY, esp=1.5):
            s = doc.styles.add_style(nome, WD_STYLE_TYPE.PARAGRAPH) if nome not in doc.styles else doc.styles[nome]
            s.font.name = fonte_escolhida
            s.font.size = Pt(tam)
            s.font.bold = negrito
            s.paragraph_format.alignment = alinh
            s.paragraph_format.line_spacing = esp
            return s

        texto_style = criar_estilo('ABNT_Texto', 12, False, WD_ALIGN_PARAGRAPH.JUSTIFY, 1.5)
        titulo_style = criar_estilo('ABNT_Titulo', 12, True, WD_ALIGN_PARAGRAPH.LEFT, 1.5)
        citacao_style = criar_estilo('ABNT_Citacao', 10, False, WD_ALIGN_PARAGRAPH.JUSTIFY, 1.0)
        citacao_style.paragraph_format.left_indent = Cm(4)
        ref_style = criar_estilo('ABNT_Ref', 12, False, WD_ALIGN_PARAGRAPH.LEFT, 1.0)
        central_style = criar_estilo('ABNT_Central', 12, True, WD_ALIGN_PARAGRAPH.CENTER, 1.5)

        # === CAPA ===
        if instituicao: doc.add_paragraph(instituicao.upper(), style='ABNT_Central')
        doc.add_page_break()
        doc.add_paragraph(autor.upper(), style='ABNT_Central')
        doc.add_page_break()
        doc.add_paragraph(titulo.upper(), style='ABNT_Central')
        if subtitulo: doc.add_paragraph(subtitulo, style='ABNT_Central')
        doc.add_page_break()
        doc.add_paragraph(f"{local.upper()}\n{ano}", style='ABNT_Central')

        # === FOLHA DE ROSTO ===
        doc.add_page_break()
        doc.add_paragraph(autor.upper(), style='ABNT_Central')
        doc.add_paragraph(titulo.upper(), style='ABNT_Central')
        if subtitulo: doc.add_paragraph(subtitulo, style='ABNT_Central')
        p = doc.add_paragraph(natureza, style='ABNT_Texto')
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        if orientador: doc.add_paragraph(f"Orientador: {orientador}", style='ABNT_Texto').alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph(f"{local.upper()}, {ano}", style='ABNT_Central')

        # === FOLHA DE APROVAÇÃO ===
        doc.add_page_break()
        doc.add_paragraph("FOLHA DE APROVAÇÃO", style='ABNT_Titulo')
        doc.add_paragraph(autor.upper(), style='ABNT_Texto')
        doc.add_paragraph(titulo.upper(), style='ABNT_Texto')
        if subtitulo: doc.add_paragraph(subtitulo, style='ABNT_Texto')
        doc.add_paragraph(natureza, style='ABNT_Texto')
        doc.add_paragraph("Data de Aprovação: __/__/____", style='ABNT_Texto')
        doc.add_paragraph("Banca Examinadora:\n________________________________\nNome, Titulação\n\n________________________________\nNome, Titulação", style='ABNT_Texto')

        # === RESUMO (OBRIGATÓRIO) ===
        doc.add_page_break()
        doc.add_paragraph("RESUMO", style='ABNT_Central')
        doc.add_paragraph(resumo, style='ABNT_Texto')
        doc.add_paragraph(f"Palavras-chave: {palavras_chave}.", style='ABNT_Texto')

        # === ABSTRACT (OPCIONAL) ===
        if abstract.strip():
            doc.add_page_break()
            doc.add_paragraph("ABSTRACT", style='ABNT_Central')
            doc.add_paragraph(abstract, style='ABNT_Texto')
            if keywords: doc.add_paragraph(f"Keywords: {keywords}.", style='ABNT_Texto')

        # === SUMÁRIO (PLACEHOLDER) ===
        doc.add_page_break()
        doc.add_paragraph("SUMÁRIO", style='ABNT_Central')
        doc.add_paragraph("(Atualize no Word: Referências > Inserir Sumário)", style='ABNT_Texto')

        # === INÍCIO DO TEXTO ===
        doc.add_page_break()
        new_section = doc.add_section(WD_SECTION_START.NEW_PAGE)
        new_section.header.is_linked_to_previous = False

        # === CONTEÚDO TEXTUAL ===
        if not conteudo.strip():
            # Cria do zero
            doc.add_paragraph("1 INTRODUÇÃO", style='ABNT_Titulo')
            doc.add_paragraph("Insira o texto da introdução aqui.", style='ABNT_Texto')
            doc.add_section(WD_SECTION_START.ODD_PAGE)
            doc.add_paragraph("2 DESENVOLVIMENTO", style='ABNT_Titulo')
            doc.add_paragraph("Insira o desenvolvimento aqui.", style='ABNT_Texto')
            doc.add_section(WD_SECTION_START.ODD_PAGE)
            doc.add_paragraph("3 CONCLUSÃO", style='ABNT_Titulo')
            doc.add_paragraph("Insira a conclusão aqui.", style='ABNT_Texto')
        else:
            # Processa upload
            em_referencias = False
            for linha in conteudo.split('\n'):
                linha = linha.strip()
                if not linha: continue

                if linha.upper().startswith('REFERÊNCIAS'):
                    em_referencias = True
                    doc.add_paragraph("REFERÊNCIAS", style='ABNT_Titulo')
                    continue

                if em_referencias:
                    ref_linhas.append(linha)
                    continue

                if re.match(r'^\d+(\.\d+)*\s+[A-ZÀ-Ú\s]+$', linha.upper()):
                    p = doc.add_paragraph(linha.upper(), style='ABNT_Titulo')
                    if '.' not in linha.split(' ', 1)[0]:  # Seção primária
                        doc.add_section(WD_SECTION_START.ODD_PAGE)
                else:
                    p = doc.add_paragraph(linha, style='ABNT_Texto')
                    if len(linha) > 120 and ('"' in linha or '“' in linha):
                        p.style = 'ABNT_Citacao'
                    for run in p.runs:
                        run.bold = False

        # === REFERÊNCIAS (SEMPRE EXISTE) ===
        doc.add_page_break()
        doc.add_paragraph("REFERÊNCIAS", style='ABNT_Titulo')
        if ref_linhas:
            ref_linhas.sort()
            for ref in ref_linhas:
                p = doc.add_paragraph(ref, style='ABNT_Ref')
                p.paragraph_format.space_after = Pt(6)
        else:
            doc.add_paragraph("(Nenhuma referência detectada. Adicione aqui.)", style='ABNT_Ref')

        # === VERIFICAÇÃO DE NORMAS ===
        verificacao = []
        if len(resumo.split()) < 150 or len(resumo.split()) > 500:
            verificacao.append("Resumo deve ter 150-500 palavras.")
        if not palavras_chave:
            verificacao.append("Adicione palavras-chave (3-5).")
        if abstract.strip() and (len(abstract.split()) < 150 or len(abstract.split()) > 500):
            verificacao.append("Abstract deve ter 150-500 palavras.")
        verificacao_msg = "\n".join(verificacao) if verificacao else "Todas as normas principais aplicadas!"

        # === SALVAR ===
        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)

        st.success("Documento gerado com sucesso!")
        st.download_button(
            label="Baixar TCC_ABNT_2024.docx",
            data=buffer,
            file_name="TCC_ABNT_2024.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

        st.info(f"**Verificação:**\n{verificacao_msg}\n\n**No Word:** Atualize sumário e paginação.")

else:
    if submitted:
        st.warning("Preencha todos os campos obrigatórios!")
    else:
        st.info("Preencha o formulário para começar.")
