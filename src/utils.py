import os
from datetime import datetime, date
from excel import get_inspections_data
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor, Inches
from unidecode import unidecode

REPORT_DIR = "./reports/"
REPORT_NAME = "RELATÓRIO MODELO"
REPORT_EXT = ".docx"

def next_filename():
    """Monta o nome do arquivo com o ID da fiscalização e caso houver arquivos já existentes incrementa em numero ao lado. EX: Relatório - ID 2 (1)"""
    inspections_data = get_inspections_data()
    id = str(int(inspections_data["ID da Fiscalização"])) 
    name = f"RELATÓRIO - ID {id}{REPORT_EXT}"
    path = os.path.join(REPORT_DIR, name)

    counter = 1
    while os.path.exists(path):
        name = f"RELATÓRIO - ID {id} ({counter}){REPORT_EXT}"
        path = os.path.join(REPORT_DIR, name)
        counter += 1
    
    return path

def search_paragraph(document, text):
    """
    Busca no documento word o texto passado como parametro e retorna a posição dele no documento
    document: arquivo (objeto)
    text: texto desejado para buscar
    """
    paragraphs_found_by_search = []
    for i, p in enumerate(document.paragraphs):
        if text in p.text:
            paragraphs_found_by_search.append(i)
    return paragraphs_found_by_search

def format_value(value):
    """
    Formata um valor para a inserção nas tabelas:
    - Retira NaN e deixa em branco
    - Converte de Float pra Int
    - Retira hora da data e deixa no padrão (DD/MM/AA)
    """
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    if isinstance(value, (datetime, date)):
        return value.strftime("%d/%m/%Y")
    return str(value)

def format_dict_values(data: dict):
    """
    Formata um dict inteiro para inserção na tabela
    """
    return {k: format_value(v) for k, v in data.items()}

def sanitize_value(value):
    """
    Converte um valor em string limpa:
    - Remove espaços extras
    - Converte para minúsculo
    - Remove acentos
    """
    return unidecode(str(value).strip().lower())

# Funções Utilitarias para outras funções, não utilizar
def replace_in_paragraph(paragraph, replacements):
        full_text = paragraph.text
        replaced_any = False
        for placeholder, value in replacements.items():
            if placeholder in full_text:
                full_text = full_text.replace(placeholder, value)
                replaced_any = True

        if replaced_any and paragraph.runs:
            first_run = paragraph.runs[0]
            first_run.text = full_text
            first_run.font.color.rgb = RGBColor(0, 0, 0)
            for run in paragraph.runs[1:]:
                run.text = ''

def set_margin(tag, value, tcMar):
        if value is not None:
            margin = tcMar.find(qn(tag))
            if margin is None:
                margin = OxmlElement(tag)
                tcMar.append(margin)
            margin.set(qn('w:w'), str(int(value * 567))) 
            margin.set(qn('w:type'), 'dxa')

# ---------------------------------------------------------------

def substitute_placeholders(document):
    """
    Define um padrão de Strings no documento ({{x}}), e substitui no documento e nas tabelas, de acordo com o dicionário retornado com os dados da fiscalização.
    Caso houver chaves iguais as Strings realiza a troca. Ex: ({{nome}}) no documento, ele percorre o dicionario e caso aja a chave nome ele troca ({{nome}})
    pelo valor correspondente a chave nome.
    """
    excel_data = get_inspections_data()

    replacements = {f"{{{{{k}}}}}": format_value(v) for k, v in excel_data.items()}

    for p in document.paragraphs:
        replace_in_paragraph(p, replacements)
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    replace_in_paragraph(p, replacements)

                
def apply_background_color(color_hex: str):
    """
    Cria um elemento de sombreado de célula com cor de fundo.
    Ex: 'D9D9D9' para cinza claro.
    """
    shading_elm = OxmlElement('w:shd')
    shading_elm.set(qn('w:val'), 'clear')
    shading_elm.set(qn('w:color'), 'auto')
    shading_elm.set(qn('w:fill'), color_hex)
    return shading_elm


def set_column_widths(table, *widths):
    """
    Ajusta a largura das colunas de uma tabela do python-docx.
    
    Parâmetros:
        table (docx.table.Table): Tabela a ser formatada.
        *widths (float): Larguras das colunas em polegadas. 
                         Passar um valor para cada coluna.
    """
    for row in table.rows:
        for col_idx, width in enumerate(widths):
            if col_idx < len(row.cells):
                row.cells[col_idx].width = Inches(width)
                
                
def set_table_margins(cell, top=None, start=None, bottom=None, end=None):
    """
    Define margens internas de cada célula de uma tabela (em Twips).
    top, start, bottom, end -> valores em cm ou None.
    """
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcMar = tcPr.find(qn('w:tcMar'))
    if tcMar is None:
        tcMar = OxmlElement('w:tcMar')
        tcPr.append(tcMar)

    set_margin('w:top', top, tcMar)
    set_margin('w:start', start, tcMar)
    set_margin('w:bottom', bottom, tcMar)
    set_margin('w:end', end, tcMar)


def set_borders_table(table):
    """
    Adiciona bordas visíveis a uma tabela do python-docx.
    Pode ser usada para qualquer tabela de documento Word.
    """
    tbl = table._element

    tblPr_list = tbl.xpath('./w:tblPr')
    if tblPr_list:
        tblPr = tblPr_list[0]
    else:
        tblPr = OxmlElement('w:tblPr')
        tbl.insert(0, tblPr)

    # Cria as bordas
    tblBorders = OxmlElement('w:tblBorders')
    for border_name in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'single')    # tipo da borda
        border.set(qn('w:sz'), '8')          # espessura
        border.set(qn('w:space'), '0')       # espaço entre borda e conteúdo
        border.set(qn('w:color'), '000000')  # cor preta
        tblBorders.append(border)

    tblPr.append(tblBorders)
    
    
def to_rows_data(data, subtitle=None):
    """
    Converte um dicionário ou lista de tuplas em lista de listas compatível com create_generic_table.
    
    Parâmetros:
    - data: dict ou lista de tuplas/listas [(chave, valor), ...]
    - subtitle: opcional, string que vira linha mesclada no topo
    
    Retorno:
    - rows_data: lista de listas pronta para create_generic_table
    """
    rows = []

    if subtitle:
        rows.append([subtitle])

    if isinstance(data, dict):
        for key, value in data.items():
            rows.append([key, value])

    elif isinstance(data, list):
        for item in data:
            rows.append(list(item))
    
    return rows


def decide_report_type():
    report_data = get_inspections_data()
    inspection_type = sanitize_value(report_data["Tipo da Fiscalização"])
    if inspection_type == "agua":
        return Document(r"./data/RELATÓRIO_AGUA_MODELO.docx")
    elif inspection_type == "esgoto":
        return Document(r"./data/RELATÓRIO_ESGOTO_MODELO.docx")
    elif inspection_type == "comercial":
        return Document(r"./data/RELATÓRIO_COMERCIAL_MODELO.docx")