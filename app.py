# app.py - REFORMATADOR ABNT 2024 (CORRIGIDO: SEM ERRO, ABSTRACT OPCIONAL, CRIAÇÃO DO ZERO)
# app.py - GERADOR ABNT 2024 (UniAmérica Padrão Perfeito)
import streamlit as st
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.section import WD_SECTION_START
import io

st.set_page_config(page_title="ABNT UniAmérica", layout="centered")
st.title("Gerador ABNT 2024 - Padrão UniAmérica")
st.caption("Preencha → Gera capa, folha de rosto e estrutura perfeita!")

# === FORMULÁRIO ===
with st.form("form_tcc"):
    curso = st.text_input("Curso (ex: CURSO DE GRADUAÇÃO EM ENGENHARIA DE SOFTWARE)", "")
    titulo = st.text_input("Título (linha 1)", "Inspeção Automatizada de Patologias: Uma abordagem de Deep Learning")
    subtitulo = st.text_input("Subtítulo (linha 2)", "para a Classificação da Criticidade de Fissuras.")
    autores = st.text_input("Nome dos alunos (separados por vírgula)", "Alessandro Rodrigues da Silva, Samuel dos Santos Gonçalves da Silva")
    orientador = st.text_input("Orientador", "Rumiê Schmoeller")
    local = st.text_input("Cidade", "Foz do Iguaçu")
    mes_ano = st.text_input("Mês e Ano", "Setembro de 2025")
    natureza = st.text_area("Natureza (para folha de rosto)", "Trabalho de Conclusão de Curso apresentado ao Curso de Graduação em Engenharia de Software da UniAmérica, como requisito parcial para obtenção do grau de Bacharel em Engenharia de Software.")
    fonte = st.selectbox("Fonte", ["Arial", "Times New Roman"])
    submitted = st.form_submit_button("Gerar Documento UniAmérica")

if submitted:
    with st.spinner("Gerando documento no padrão UniAmérica..."):
        doc = Document()

        # === ESTILOS ===
        def estilo(nome, tam=12, negrito=False, alinh=WD_ALIGN_PARAGRAPH.CENTER, esp=1.5):
            s = doc.styles.add_style(nome, WD_STYLE_TYPE.PARAGRAPH) if nome not in doc.styles else doc.styles[nome]
            s.font.name = fonte
            s.font.size = Pt(tam)
            s.font.bold = negrito
            s.paragraph_format.alignment = alinh
            s.paragraph_format.line_spacing = esp
            s.paragraph_format.space_after = Pt(0)
            return s

        central = estilo('Central', 12, False, WD_ALIGN_PARAGRAPH.CENTER, 1.5)
        central_bold = estilo('CentralBold', 12, True, WD_ALIGN_PARAGRAPH.CENTER, 1.5)
        direita = estilo('Direita', 12, False, WD_ALIGN_PARAGRAPH.RIGHT, 1.5)

        # === FUNÇÃO: NOVA PÁGINA COM MARGENS ABNT ===
        def nova_pagina():
            doc.add_section(WD_SECTION_START.NEW_PAGE)
            sec = doc.sections[-1]
            sec.top_margin = Cm(3)
            sec.bottom_margin = Cm(2)
            sec.left_margin = Cm(3)
            sec.right_margin = Cm(2)
            return sec

        # === CAPA ===
        # Logo (opcional - você pode subir depois no Word)
        doc.add_paragraph("")  # espaço para logo

        # Curso
        p = doc.add_paragraph(curso.upper(), style='CentralBold')
        p.paragraph_format.space_after = Pt(60)

        # Título
        p = doc.add_paragraph(titulo, style='Central')
        p.paragraph_format.space_after = Pt(12)
        doc.add_paragraph(subtitulo, style='Central')

        # Espaço grande
        p = doc.add_paragraph("", style='Central')
        p.paragraph_format.space_after = Pt(120)

        # Alunos
        doc.add_paragraph(f"Nome dos alunos: {autores}", style='Central')
        doc.add_paragraph(f"Orientador: {orientador}", style='Central')

        # Local e data
        p = doc.add_paragraph("", style='Central')
        p.paragraph_format.space_after = Pt(80)
        doc.add_paragraph(f"{local}, {mes_ano}", style='Central')

        # === FOLHA DE ROSTO ===
        nova_pagina()

        # Logo (espaço)
        doc.add_paragraph("")

        # Curso
        p = doc.add_paragraph(curso.upper(), style='CentralBold')
        p.paragraph_format.space_after = Pt(60)

        # Título
        doc.add_paragraph(titulo, style='Central')
        doc.add_paragraph(subtitulo, style='Central')

        # Espaço
        p = doc.add_paragraph("", style='Central')
        p.paragraph_format.space_after = Pt(80)

        # Alunos
        doc.add_paragraph(f"Nome dos alunos: {autores}", style='Central')
        doc.add_paragraph(f"Orientador: {orientador}", style='Central')

        # Natureza (à direita)
        p = doc.add_paragraph(natureza, style='Direita')
        p.paragraph_format.space_before = Pt(60)

        # Local e data
        p = doc.add_paragraph(f"{local}, {mes_ano}", style='Central')
        p.paragraph_format.space_before = Pt(60)

        # === PRÓXIMAS PÁGINAS (opcional) ===
        nova_pagina()
        doc.add_paragraph("FOLHA DE APROVAÇÃO", style='CentralBold')
        nova_pagina()
        doc.add_paragraph("RESUMO", style='CentralBold')
        nova_pagina()
        doc.add_paragraph("SUMÁRIO", style='CentralBold')
        nova_pagina()
        doc.add_paragraph("1 INTRODUÇÃO", style='CentralBold')

        # === SALVAR ===
        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)

        st.success("Documento gerado no padrão UniAmérica!")
        st.download_button(
            label="Baixar TCC_UniAmerica_ABNT2024.docx",
            data=buffer,
            file_name="TCC_UniAmerica_ABNT2024.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

        st.info("""
        **Próximos passos no Word:**
        1. Insira o **logo da UniAmérica** no topo da capa
        2. Atualize o **sumário** (Referências > Inserir Sumário)
        3. Adicione **numeração de página** a partir da Introdução
        4. Salve como PDF
        """)
