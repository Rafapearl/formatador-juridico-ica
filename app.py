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
                                 cor_texto=None, recuo_lista=False, recuo_primeira_linha=True):
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
    elif recuo_primeira_linha:
        paragrafo.paragraph_format.first_line_indent = Cm(1.27)

    for run in paragrafo.runs:
        run.font.name = 'Arial'
        run.font.size = Pt(tamanho_fonte)
        run.bold = negrito
        run.italic = italico
        if cor_texto:
            run.font.color.rgb = RGBColor(*cor_texto)

def detectar_tipo_paragrafo(texto):
    texto_limpo = texto.strip()

    if re.match(r'^\s*(PEDIDOS|POR TUDO ISSO)', texto_limpo, re.IGNORECASE):
        return 'secao_pedidos', True, 'center'
        
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

def formatar_documento(doc_entrada, doc_saida_path, logo_path=None, debug_mode=False):
    debug_info = []
    doc_novo = Document()

    for section in doc_novo.sections:
        section.top_margin = Cm(2.5)
        section.bottom_margin = Cm(2.5)
        section.left_margin = Cm(3)
        section.right_margin = Cm(2)

    criar_cabecalho(doc_novo, logo_path)
    em_pedidos = False  # Variável de controle corrigida

    for i, para in enumerate(doc_entrada.paragraphs):
        texto = para.text.strip()

        if not texto:
            doc_novo.add_paragraph()
            continue

        tipo, negrito, alinhamento = detectar_tipo_paragrafo(texto)
        
        # Controle da seção de Pedidos
        if tipo == 'secao_pedidos':
            em_pedidos = True
        elif tipo in ['secao_principal', 'cabecalho', 'titulo_acao']:
            em_pedidos = False

        # Ajustar formatação se estiver em Pedidos
        if em_pedidos and tipo not in ['secao_pedidos']:
            alinhamento = 'center'
            recuo_pedidos = False
        else:
            recuo_pedidos = True

        if debug_mode:
            debug_info.append({
                "index": i,
                "texto": texto[:50] + "..." if len(texto) > 50 else texto,
                "tipo_detectado": tipo,
                "negrito": negrito,
                "alinhamento": alinhamento
            })

        p = doc_novo.add_paragraph()
        run = p.add_run(texto)

        if tipo == 'cabecalho':
            aplicar_formatacao_paragrafo(p, alinhamento='center', negrito=True,
                               tamanho_fonte=12, espacamento_antes=0,
                               espacamento_depois=40, recuo_primeira_linha=False)
        
        elif tipo == 'titulo_acao':
            aplicar_formatacao_paragrafo(p, alinhamento='center', negrito=True,
                               tamanho_fonte=12, espacamento_antes=30,
                               espacamento_depois=24,
                               cor_texto=FORMATO_CONFIG['cor_titulo'],
                               recuo_primeira_linha=False)
        
        elif tipo == 'secao_principal':
            aplicar_formatacao_paragrafo(p, alinhamento='left', negrito=True,
                               tamanho_fonte=12, espacamento_antes=12,
                               espacamento_depois=6,
                               cor_texto=FORMATO_CONFIG['cor_secao'],
                               recuo_primeira_linha=False)
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
            
        elif tipo == 'secao_pedidos':
            aplicar_formatacao_paragrafo(p, 
                alinhamento='center',
                negrito=True,
                tamanho_fonte=12,
                espacamento_antes=24,
                espacamento_depois=12,
                recuo_primeira_linha=False
            )
            adicionar_linha_horizontal(p, FORMATO_CONFIG['cor_linha'])
        
        else:
            aplicar_formatacao_paragrafo(p, 
                alinhamento=alinhamento, 
                negrito=negrito,
                tamanho_fonte=12,
                espacamento_antes=6,
                espacamento_depois=6,
                recuo_primeira_linha=recuo_pedidos
            )

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
    footer = last_section.footer

    for p in footer.paragraphs:
        if p.text.strip():
            p_element = p._element.get_or_add_pPr()
            from docx.oxml import parse_xml
            frame_xml = parse_xml(
            '<w:framePr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
            'w:w="13000" w:h="2500" w:wrap="around" w:vAnchor="page" w:hAnchor="page" w:xAlign="center" />'
             )
            p_element.append(frame_xml)
            p.paragraph_format.space_before = Pt(50)
            p.paragraph_format.space_after = Pt(50)

    doc_novo.save(doc_saida_path)
    
    if debug_mode:
        return doc_saida_path, debug_info
    else:
        return doc_saida_path

# ... (O restante das funções criar_arquivo_zip e main permanecem inalteradas)

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

    if 'debug_mode' not in st.session_state:
        st.session_state.debug_mode = False

    st.title("📄 Formatador Jurídico ICA Advocacia")
    st.write("Ferramenta para formatação automática de documentos jurídicos.")
    st.markdown("---")

    with st.sidebar:
        st.header("⚙️ Configurações")
        save_logo = st.checkbox("Salvar logo para uso futuro", value=True)        
        st.session_state.debug_mode = st.checkbox("Modo de depuração", value=False)
        if st.session_state.debug_mode:
            st.info("O modo de depuração mostrará informações detalhadas sobre a formatação.")
    
    # Seções de upload e processamento permanecem iguais
    # ... (O restante do código da função main permanece inalterado)

if __name__ == "__main__":
    main()
