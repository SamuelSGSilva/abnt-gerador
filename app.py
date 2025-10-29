# app.py - GERADOR ABNT 2024 WEB
import streamlit as st
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
import io
import re

st.set_page_config(page_title="Gerador ABNT 2024", layout="centered")
st.title("🚀 Gerador ABNT 2024 Automático")
st.caption("Conforme NBR 14724:2024 — Zero formatação manual! Cole seu texto abaixo.")

# Entrada
with st.expander("📝 Exemplo de Texto de Entrada (copie e cole)"):
    st.code("""
@TITULO: Seu Título Aqui
@AUTOR: Seu Nome
@RESUMO: Seu resumo...
@SECAO: 1 INTRODUÇÃO
Seu texto...
@REFERENCIAS:
Autor. Título. Ano.
    """, language="text")

texto = st.text_area("Cole seu conteúdo aqui (use @TITULO, @SECAO, etc.)", height=300, placeholder="Ex: @TITULO: Meu TCC")
arquivo = st.file_uploader("Ou faça upload do .txt", type="txt")

if arquivo:
    texto = arquivo.read().decode("utf-8")

if st.button("🔥 Gerar Documento ABNT 2024") and texto:
    with st.spinner("Gerando documento perfeito..."):
        # === PROCESSAR TEXTO ===
        dados = {}
        secoes = []
        referencias = []
        modo = None
        linhas = texto.split('\n')

        for linha in linhas:
            linha = linha.strip()
            if linha.startswith('@TITULO:'): dados['titulo'] = linha[9:].strip()
            elif linha.startswith('@SUBTITULO:'): dados['subtitulo'] = linha[11:].strip()
            elif linha.startswith('@AUTOR:'): dados['autor'] = linha[7:].strip()
            elif linha.startswith('@ORIENTADOR:'): dados['orientador'] = linha[12:].strip()
            elif linha.startswith('@INSTITUICAO:'): dados['instituicao'] = linha[13:].strip()
            elif linha.startswith('@LOCAL:'): dados['local'] = linha[7:].strip()
            elif linha.startswith('@ANO:'): dados['ano'] = linha[5:].strip()
            elif linha.startswith('@RESUMO:'): modo = 'resumo'; dados['resumo'] = ''
            elif linha.startswith('@ABSTRACT:'): modo = 'abstract'; dados['abstract'] = ''
            elif linha.startswith('@PALAVRAS_CHAVE:'): dados['palavras'] = linha[16:].strip()
            elif linha.startswith('@KEYWORDS:'): dados['keywords'] = linha[11:].strip()
            elif linha.startswith('@SECAO:'):
                if secoes and modo == 'secoes': secoes[-1]['texto'] = secoes[-1]['texto'].strip()
                secoes.append({'titulo': linha[8:].strip(), 'texto': ''})
                modo = 'secoes'
            elif linha.startswith('@REFERENCIAS:'): modo = 'referencias'
            elif linha and modo == 'resumo': dados['resumo'] += linha + ' '
            elif linha and modo == 'abstract': dados['abstract'] += linha + ' '
            elif linha and modo == 'secoes' and secoes: secoes[-1]['texto'] += linha + ' '
            elif linha and modo == 'referencias': referencias.append(linha)

        # Validação básica
        if not dados.get('titulo'):
            st.error("❌ Adicione @TITULO: no texto!")
            st.stop()

        # === CRIAR DOC ===
        doc = Document()
        section = doc.sections[0]
        section.top_margin = Cm(3); section.bottom_margin = Cm(2)
        section.left_margin = Cm(3); section.right_margin = Cm(2)

        def estilo(nome, fonte='Arial', tam=12, negrito=False, alinhamento=WD_ALIGN_PARAGRAPH.JUSTIFY):
            s = doc.styles.add_style(nome, WD_STYLE_TYPE.PARAGRAPH)
            f = s.font
            f.name = fonte; f.size = Pt(tam); f.bold = negrito
            s.paragraph_format.line_spacing = 1.5 if tam == 12 else 1.0
            s.paragraph_format.alignment = alinhamento
            return s

        estilo('Capa', 'Arial', 12, True, WD_ALIGN_PARAGRAPH.CENTER)
        estilo('Texto', 'Arial', 12, False, WD_ALIGN_PARAGRAPH.JUSTIFY)
        estilo('Secao', 'Arial', 12, True, WD_ALIGN_PARAGRAPH.LEFT)

        def add(texto, estilo='Texto', central=False):
            par = doc.add_paragraph(texto, style=estilo)
            if central: par.alignment = WD_ALIGN_PARAGRAPH.CENTER
            return par

        # Capa
        if dados.get('instituicao'): add(dados['instituicao'], 'Capa', True)
        doc.add_page_break()
        add(dados.get('autor', ''), 'Capa', True)
        doc.add_page_break()
        add(dados.get('titulo', ''), 'Capa', True)
        if dados.get('subtitulo'): add(dados['subtitulo'], 'Capa', True)
        doc.add_page_break()
        add(dados.get('local', '') + ', ' + dados.get('ano', ''), 'Capa', True)

        # Folha de rosto
        doc.add_page_break()
        add(dados.get('autor', ''), 'Capa', True)
        add(dados.get('titulo', ''), 'Capa', True)
        if dados.get('subtitulo'): add(dados['subtitulo'], 'Capa', True)
        add(f"Orientador: {dados.get('orientador', '')}", 'Texto', True)
        add(dados.get('local', '') + ', ' + dados.get('ano', ''), 'Capa', True)

        # Resumo
        doc.add_page_break()
        add('RESUMO', 'Secao')
        add(dados.get('resumo', '').strip(), 'Texto')
        add(f"Palavras-chave: {dados.get('palavras', '')}", 'Texto')

        # Abstract
        doc.add_page_break()
        add('ABSTRACT', 'Secao')
        add(dados.get('abstract', '').strip(), 'Texto')
        add(f"Keywords: {dados.get('keywords', '')}", 'Texto')

        # Seções
        for s in secoes:
            doc.add_page_break()
            add(s['titulo'], 'Secao')
            add(s['texto'].strip(), 'Texto')

        # Referências
        doc.add_page_break()
        add('REFERÊNCIAS', 'Secao')
        ref_style = doc.styles.add_style('Ref', WD_STYLE_TYPE.PARAGRAPH)
        ref_style.font.size = Pt(12)
        ref_style.paragraph_format.line_spacing = 1.0
        for r in referencias:
            doc.add_paragraph(r, style='Ref')

        # Salvar em memória
        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)

        st.success("✅ Documento ABNT 2024 gerado com sucesso!")
        st.download_button(
            label="📥 Baixar trabalho_final_abnt.docx",
            data=buffer,
            file_name="trabalho_final_abnt.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
else:
    st.info("💡 Dica: Use o expander acima para um exemplo pronto!")
