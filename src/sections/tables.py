from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH
from excel import get_non_conformities
from utils import search_paragraph, apply_background_color
from images import set_borders_table  

def create_non_conformities_table(document: Document):
    df_ncs = get_non_conformities()

    selected_columns = [
        "Unidade", 
        "Não Conformidade", 
        "Nome da Foto", 
        "Artigo", 
        "Enquadramento", 
        "Determinações"
    ]
    df_ncs = df_ncs[selected_columns]

    table = document.add_table(rows=1, cols=len(selected_columns))
    hdr_cells = table.rows[0].cells

    # Cabeçalho estilizado
    for idx, col_name in enumerate(selected_columns):
        cell = hdr_cells[idx]

        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

        paragraph = cell.paragraphs[0]
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = paragraph.add_run(col_name.upper()) 
        run.bold = True
        run.font.name = 'Arial'
        run.font.size = Pt(8)

        # Fundo cinza
        cell._element.get_or_add_tcPr().append(apply_background_color("D9D9D9"))

    # Dados
    for _, row in df_ncs.iterrows():
        cells = table.add_row().cells
        for idx, col_name in enumerate(selected_columns):
            cell = cells[idx]

            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

            paragraph = cell.paragraphs[0]
            paragraph.style = 'Arial10'
            run = paragraph.add_run(str(row[col_name]))
            run.font.name = 'Arial'
            run.font.size = Pt(8)

    # Inserir após parágrafo
    list_paragraph_index = search_paragraph(document, "Lista de NCs do SAA {{Municipio}}")[0]
    list_paragraph = document.paragraphs[list_paragraph_index]
    list_paragraph._element.addnext(table._element)

    # Aplica bordas
    set_borders_table(table)

    print("✅ Tabela criada com cabeçalho estilizado, fonte Arial10 e bordas.")


