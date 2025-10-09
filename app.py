import streamlit as st
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import re
import os
import tempfile
import zipfile
import io
from datetime import datetime

# Configurações do rodapé
RODAPE_CONFIG = {
    'endereco': 'Avenida Cristovão Colombo, nº 485, 4º andar, Savassi, Belo Horizonte/MG',
    'telefone': '(31) 9 9703-9242',
    'email': 'contato@icaadvocacia.com.br',
    'cor_fundo': (0, 0, 102),
    'cor_texto': (255, 255, 255),
    'largura': '150%',
    'altura': '150%'
}

# Configurações de formatação
FORMATO_CONFIG = {
    'fonte_padrao': 'Arial',
    'tamanho_fonte_normal': 12,
    'tamanho_fonte_titulo': 12,
    'cor_titulo': (59, 75, 160),
    'cor_secao': (59, 75, 160),
    'cor_linha': (192, 192, 192),
    'espacamento_antes': Pt(6),
    'espacamento_depois': Pt(6),
    'espacamento_linha': 1.5
}

def criar_cabecalho(doc, logo_path=None):
    section = doc.sections[0]
    header = section.header

    for para in header.paragraphs:
        para.clear()

    p = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    if logo_path and os.path.exists(logo_path):
        run = p.add_run()
        run.add_picture(logo_path, width=Inches(2.5))

    p.paragraph_format.space_after = Pt(24)
    p.paragraph_format.space_before = Pt(12)

def criar_rodape(doc, config):
    section = doc.sections[0]
    footer = section.footer

    for para in footer.paragraphs:
        para.clear()

    p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    shading_elm = OxmlElement('w:shd')
    shading_elm.set(qn('w:fill'), '{:02x}{:02x}{:02x}'.format(*config['cor_fundo']))
    p._element.get_or_add_pPr().append(shading_elm)

    p.paragraph_format.space_before = Pt(50)
    p.paragraph_format.space_after = Pt(50)

    run_space_before = p.add_run("\n\n")
    run_space_before.font.size = Pt(2)

    run1 = p.add_run(config['endereco'])
    run1.font.color.rgb = RGBColor(*config['cor_texto'])
    run1.font.size = Pt(10)
    run1.font.name = 'Arial'

    p.add_run('\n')
    run2 = p.add_run(f"{config['telefone']} | {config['email']}")
    run2.font.color.rgb = RGBColor(*config['cor_texto'])
    run2.font.size = Pt(10)
    run2.font.name = 'Arial'
    p.add_run("\n\n\n")

def adicionar_linha_horizontal(paragrafo, cor_rgb=(192, 192, 192)):
    p = paragrafo._element
    pPr = p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '6')
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), '{:02x}{:02x}{:02x}'.format(*cor_rgb))
    pBdr.append(bottom)
    pPr.append(pBdr)

def aplicar_formatacao_paragrafo(paragrafo, alinhamento='justify', negrito=False,
                                 italico=False, tamanho_fonte=12, espacamento_antes=6,
                                 espacamento_depois=6, espacamento_linha=1.5, 
                                 cor_texto=None, recuo_lista=False):
    if alinhamento == 'center':
        paragrafo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif alinhamento == 'justify':
        paragrafo.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    elif alinhamento == 'left':
        paragrafo.alignment = WD_ALIGN_PARAGRAPH.LEFT

    paragrafo.paragraph_format.space_before = Pt(espacamento_antes)
    paragrafo.paragraph_format.space_after = Pt(espacamento_depois)
    paragrafo.paragraph_format.line_spacing = espacamento_linha

    if recuo_lista:
        paragrafo.paragraph_format.left_indent = Inches(0.25)
        paragrafo.paragraph_format.first_line_indent = Inches(-0.25)

    for run in paragrafo.runs:
        run.font.name = 'Arial'
        run.font.size = Pt(tamanho_fonte)
        run.bold = negrito
        run.italic = italico
        if cor_texto:
            run.font.color.rgb = RGBColor(*cor_texto)

def detectar_tipo_paragrafo(texto):
    texto_limpo = texto.strip()

    if re.match(r'^\s*Doc\.\s*\d+', texto_limpo):
        return 'item_doc', False, 'left'

    if re.match(r'^(EXMO|EXCELENTÍSSIM[OA])\b', texto_limpo, re.IGNORECASE):
        return 'cabecalho', True, 'center'

    aspas = ['"', "'", '“', '”', '‘', '’', '«', '»', '‹', '›', '„', '‟', '「', '」', '『', '』']
    if any(aspas in texto_limpo for aspas in aspas):
        return 'citacao', False, 'justify'
    
    if (re.match(r'Art\.\s*\d+', texto_limpo) or
        re.match(r'§\s*\d+', texto_limpo) or
        re.search(r'inciso\s+[IVX]+', texto_limpo) or
        re.search(r'alínea\s+[a-z]', texto_limpo)):
        return 'citacao', False, 'justify'
    
    if re.match(r'^\s*[•▪■□◊○●◉◎◌◦⦿⦾]+\s+', texto_limpo):
        return 'subsecao', True, 'left'
    
    if re.match(r'^\s*\d+[\.\)]\s+', texto_limpo):
        return 'lista', False, 'left'
    
    if re.match(r'^\s*[a-z][\.\)]\s+', texto_limpo):
        return 'lista', False, 'left'
    
    if re.match(r'^[IVX]+[\s]*[.–—\-]+[\s]*(DOS?|DAS?)[\s]+[A-ZÀÁÂÃÉÊÍÓÔÕÚÇ\s]+$', texto_limpo):
        return 'secao_principal', True, 'left'
    
    if ('AÇÃO DE' in texto_limpo.upper() and 
        len(texto_limpo) < 150 and 
        texto_limpo.upper().count(' ') >= 2 and
        not any(marcador in texto_limpo for marcador in ['•', '▪', '-', '*']) and
        not re.match(r'^\s*\d+[\.\)]', texto_limpo)):
        return 'titulo_acao', True, 'center'
    
    if re.match(r'^\s*[\-–—*+]\s+', texto_limpo):
        return 'lista', False, 'left'
    
    words = texto_limpo.split()
    if (2 <= len(words) <= 7 and 
        all(w[0].isupper() for w in words if len(w) > 3) and 
        not texto_limpo.endswith('.') and
        len(texto_limpo) < 50):
        return 'normal', True, 'left'
    
    return 'normal', False, 'justify'

def formatar_documento(doc_entrada, doc_saida_path, logo_path=None):
    doc_novo = Document()

    for section in doc_novo.sections:
        section.top_margin = Cm(2.5)
        section.bottom_margin = Cm(2.5)
        section.left_margin = Cm(3)
        section.right_margin = Cm(2)

    criar_cabecalho(doc_novo, logo_path)

    for para in doc_entrada.paragraphs:
        texto = para.text.strip()

        if not texto:
            doc_novo.add_paragraph()
            continue

        tipo, negrito, alinhamento = detectar_tipo_paragrafo(texto)
        p = doc_novo.add_paragraph()
        run = p.add_run(texto)
        
        if tipo == 'cabecalho':
            aplicar_formatacao_paragrafo(p, alinhamento='center', negrito=True,
                                       tamanho_fonte=12, espacamento_antes=0,
                                       espacamento_depois=40)
        
        elif tipo == 'titulo_acao':
            aplicar_formatacao_paragrafo(p, alinhamento='center', negrito=True,
                                       tamanho_fonte=12, espacamento_antes=30,
                                       espacamento_depois=24,
                                       cor_texto=FORMATO_CONFIG['cor_titulo'])
        
        elif tipo == 'secao_principal':
            aplicar_formatacao_paragrafo(p, alinhamento='left', negrito=True,
                                       tamanho_fonte=12, espacamento_antes=12,
                                       espacamento_depois=6,
                                       cor_texto=FORMATO_CONFIG['cor_secao'])
            adicionar_linha_horizontal(p, FORMATO_CONFIG['cor_linha'])
        
        elif tipo == 'item_doc':
            aplicar_formatacao_paragrafo(p, alinhamento='left', negrito=False,
                              tamanho_fonte=12, espacamento_antes=6,
                              espacamento_depois=6, recuo_lista=True)
        
        elif tipo == 'subsecao':
            aplicar_formatacao_paragrafo(p, alinhamento='left', negrito=True,
                              tamanho_fonte=12, espacamento_antes=6,
                              espacamento_depois=6, recuo_lista=True)
        
        elif tipo == 'citacao':
            aplicar_formatacao_paragrafo(p, alinhamento='justify', negrito=False,
                              italico=True, tamanho_fonte=11,
                              espacamento_antes=6, espacamento_depois=6,
                              recuo_lista=True)
        
        elif tipo == 'lista':
            aplicar_formatacao_paragrafo(p, alinhamento='left', negrito=False,
                              tamanho_fonte=12, espacamento_antes=3,
                              espacamento_depois=3, recuo_lista=True)
        
        else:
            aplicar_formatacao_paragrafo(p, alinhamento='justify', negrito=False,
                                       tamanho_fonte=12, espacamento_antes=6,
                                       espacamento_depois=6)

    for table in doc_entrada.tables:
        nova_tabela = doc_novo.add_table(rows=len(table.rows), cols=len(table.columns))
        nova_tabela.style = 'Light Grid Accent 1'

        for i, row in enumerate(table.rows):
            for j, cell in enumerate(row.cells):
                nova_tabela.rows[i].cells[j].text = cell.text
                if i == 0:
                    for paragraph in nova_tabela.rows[i].cells[j].paragraphs:
                        for run in paragraph.runs:
                            run.font.bold = True
                            run.font.size = Pt(11)

        doc_novo.add_paragraph()

    last_section = doc_novo.sections[-1]
    criar_rodape(doc_novo, RODAPE_CONFIG)

    footer = last_section.footer
    for p in footer.paragraphs:
        if p.text.strip():
            p_element = p._element.get_or_add_pPr()
            frame_xml = parse_xml(
            '<w:framePr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
            'w:w="13000" w:h="2500" w:wrap="around" w:vAnchor="page" w:hAnchor="page" w:xAlign="center" />'
            )
            p_element.append(frame_xml)
            p.paragraph_format.space_before = Pt(50)
            p.paragraph_format.space_after = Pt(50)

    doc_novo.save(doc_saida_path)
    return doc_saida_path

def criar_arquivo_zip(arquivos):
    zip_buffer = io.BytesIO()
    
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        for arquivo in arquivos:
            zip_file.write(arquivo, os.path.basename(arquivo))
    
    zip_buffer.seek(0)
    return zip_buffer

def main():
    st.set_page_config(
        page_title="Formatador Jurídico ICA Advocacia", 
        page_icon="📄",
        layout="wide"
    )

    st.title("📄 Formatador Jurídico ICA Advocacia")
    st.write("Ferramenta para formatação automática de documentos jurídicos.")
    st.markdown("---")

    with st.sidebar:
        st.header("⚙️ Configurações")
        save_logo = st.checkbox("Salvar logo para uso futuro", value=True)

    st.header("1️⃣ Upload do Logo ICA")
    logo_container = st.container()
    with logo_container:
        logo_col1, logo_col2 = st.columns([2, 1])
        
        with logo_col1:
            logo_file = st.file_uploader("Faça upload do logo ICA (arquivo PNG, JPG)", 
                                        type=["png", "jpg", "jpeg"], 
                                        key="logo")

        with logo_col2:
            if 'logo_cache' in st.session_state and logo_file is None:
                logo_path = st.session_state.logo_cache
                st.success("✅ Logo salvo em uso")
                st.image(logo_path, width=150)
            elif logo_file is not None:
                temp_logo = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
                temp_logo.write(logo_file.getvalue())
                logo_path = temp_logo.name

                if save_logo:
                    st.session_state.logo_cache = logo_path
                
                st.success("✅ Logo carregado!")
                st.image(logo_file, width=150)
            else:
                logo_path = None
                st.info("Logo não carregado. O cabeçalho será criado sem logo.")
    
    st.markdown("---")
    st.header("2️⃣ Upload dos Documentos Word")
    
    doc_container = st.container()
    with doc_container:
        uploaded_files = st.file_uploader(
            "Selecione os documentos Word (.docx) para formatação",
            type=["docx"], 
            accept_multiple_files=True,
            key="docs"
        )

        if not uploaded_files:
            st.warning("⚠️ Nenhum documento carregado ainda.")
        else:
            st.success(f"✅ {len(uploaded_files)} documento(s) carregado(s)")
            with st.expander(f"📋 Ver arquivos carregados ({len(uploaded_files)})"):
                for i, doc_file in enumerate(uploaded_files):
                    st.write(f"{i+1}. {doc_file.name}")
    
    st.markdown("---")
    st.header("3️⃣ Formatação")
    
    format_button = st.button(
        "Formatar Documentos", 
        disabled=(len(uploaded_files) == 0 or (logo_path is None and 'logo_cache' not in st.session_state)),
        type="primary",
        use_container_width=True
    )

    if format_button and len(uploaded_files) > 0:
        if logo_path is None and 'logo_cache' in st.session_state:
            logo_path = st.session_state.logo_cache
            
        progress_container = st.container()
        
        with progress_container:
            status_text = st.empty()
            progress_bar = st.progress(0)
            
            temp_dir = tempfile.mkdtemp()
            processed_count = st.empty()
            processed_count.info("Preparando processamento...")
            
            arquivos_processados = []
            errors = []
            
            for i, doc_file in enumerate(uploaded_files):
                try:
                    status_text.info(f"⏳ Processando: {doc_file.name} ({i+1}/{len(uploaded_files)})")
                    
                    input_path = os.path.join(temp_dir, doc_file.name)
                    with open(input_path, "wb") as f:
                        f.write(doc_file.getvalue())
                    
                    nome_base = os.path.splitext(doc_file.name)[0]
                    output_path = os.path.join(temp_dir, f"{nome_base}_FORMATADO.docx")
                    
                    doc = Document(input_path)
                    output_path = formatar_documento(doc, output_path, logo_path)
                    arquivos_processados.append(output_path)
                    
                    progress = int(((i + 1) / len(uploaded_files)) * 100)
                    progress_bar.progress(progress)
                    processed_count.info(f"✅ Processados: {i+1}/{len(uploaded_files)} documentos")
                    
                except Exception as e:
                    errors.append((doc_file.name, str(e)))
            
            if arquivos_processados:
                status_text.success(f"✅ Processamento concluído! {len(arquivos_processados)} documento(s) formatado(s).")
                progress_bar.progress(100)
                
                if errors:
                    with st.expander(f"⚠️ Erros ({len(errors)})"):
                        for file_name, error_msg in errors:
                            st.error(f"Arquivo: {file_name} - Erro: {error_msg}")
                
                st.markdown("---")
                st.header("4️⃣ Download dos Documentos Formatados")
                
                data_atual = datetime.now().strftime("%Y%m%d")
                zip_filename = f"Documentos_Formatados_ICA_{data_atual}.zip"
                
                if len(arquivos_processados) > 1:
                    st.subheader("📦 Download de todos os arquivos")
                    zip_data = criar_arquivo_zip(arquivos_processados)
                    
                    col1, col2 = st.columns([3, 1])
                    with col1:
                        st.download_button(
                            label="⬇️ BAIXAR TODOS OS DOCUMENTOS DE UMA VEZ",
                            data=zip_data,
                            file_name=zip_filename,
                            mime="application/zip",
                            use_container_width=True,
                            type="primary",
                        )
                    
                    with col2:
                        st.info(f"{len(arquivos_processados)} arquivos no ZIP")
                    
                    st.markdown("---")
                
                st.subheader("📄 Downloads individuais")
                num_cols = 2
                file_cols = st.columns(num_cols)
                
                for i, file_path in enumerate(arquivos_processados):
                    col_idx = i % num_cols
                    with file_cols[col_idx]:
                        file_name = os.path.basename(file_path)
                        with open(file_path, "rb") as file:
                            st.download_button(
                                label=f"⬇️ {file_name}",
                                data=file,
                                file_name=file_name,
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                key=f"download_{i}",
                                use_container_width=True
                            )
                        st.write("")
            else:
                st.error("❌ Nenhum documento foi processado com sucesso.")

    st.markdown("---")
    with st.expander("ℹ️ Informações sobre o Formatador"):
        st.markdown("""
        ### Sobre a Formatação
        - **Cabeçalho**: Logo centralizado
        - **Títulos**: Em azul com formatação adequada
        - **Seções principais**: Com linha horizontal
        - **Rodapé**: Em azul escuro com informações do escritório
        
        ### Dicas de Uso
        - Processe vários documentos de uma vez
        - Use nomes descritivos para seus arquivos
        - O logo será mantido para uso futuro
        """)
    
    st.markdown("---")
    st.markdown("<div style='text-align: center; color: gray; font-size: 0.8em;'>Formatador Jurídico ICA Advocacia - Versão 1.0</div>", unsafe_allow_html=True)

if __name__ == "__main__":
    main()
