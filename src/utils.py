import os
from datetime import datetime, date
from excel import get_inspections_data
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor, Inches

REPORT_DIR = "./reports/"
REPORT_NAME = "RELATÓRIO MODELO"
REPORT_EXT = ".docx"

def next_filename():
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
    paragraphs_found_by_search = []
    print(">>> Busca no documento iniciada")
    for i, p in enumerate(document.paragraphs):
        if text in p.text:
            paragraphs_found_by_search.append(i)
            print(">>> Uma ocorrencia foi encontrada")
    return paragraphs_found_by_search

def format_value(value):
        if value is None:
            return ""
        if isinstance(value, float) and value.is_integer():
            return str(int(value))
        if isinstance(value, (datetime, date)):
            return value.strftime("%d/%m/%Y")
        return str(value)

def format_dict_values(data: dict):
    return {k: format_value(v) for k, v in data.items()}

def substitute_placeholders(document):
    excel_data = get_inspections_data()

    replacements = {f"{{{{{k}}}}}": format_value(v) for k, v in excel_data.items()}

    def replace_in_paragraph(paragraph):
        full_text = paragraph.text
        replaced_any = False
        for placeholder, value in replacements.items():
            if placeholder in full_text:
                full_text = full_text.replace(placeholder, value)
                replaced_any = True

        if replaced_any and paragraph.runs:
            # Mantém formatação do primeiro run
            first_run = paragraph.runs[0]
            first_run.text = full_text
            first_run.font.color.rgb = RGBColor(0, 0, 0)
            # Limpa os demais runs
            for run in paragraph.runs[1:]:
                run.text = ''

    # Aplica nos parágrafos
    for p in document.paragraphs:
        replace_in_paragraph(p)

    # Aplica nas tabelas
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    replace_in_paragraph(p)


                
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

    def set_margin(tag, value):
        if value is not None:
            margin = tcMar.find(qn(tag))
            if margin is None:
                margin = OxmlElement(tag)
                tcMar.append(margin)
            margin.set(qn('w:w'), str(int(value * 567)))  # 1 cm ≈ 567 twips
            margin.set(qn('w:type'), 'dxa')

    set_margin('w:top', top)
    set_margin('w:start', start)
    set_margin('w:bottom', bottom)
    set_margin('w:end', end)
