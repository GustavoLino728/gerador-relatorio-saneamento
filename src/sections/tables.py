from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH
from unidecode import unidecode
from excel import get_non_conformities, get_inspections_data, list_of_all_units, documents_excel, town_statistics
from utils import search_paragraph, apply_background_color
from images import set_borders_table  
import pandas as pd

def create_non_conformities_table(document: Document, text):
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
    list_paragraph_index = search_paragraph(document, text)[0]
    list_paragraph = document.paragraphs[list_paragraph_index]
    list_paragraph._element.addnext(table._element)

    # Aplica bordas
    set_borders_table(table)

    print("✅ Tabela criada com cabeçalho estilizado, fonte Arial10 e bordas.")

def create_town_units_table(document, text):
    report_data = get_inspections_data()

    # Normaliza dados
    town_name = unidecode(str(report_data["Municipio"]).strip().lower())
    inspection_type = unidecode(str(report_data["Tipo da Fiscalização"]).strip().lower())

    units_df = list_of_all_units.copy()
    units_df.columns = units_df.columns.str.strip()

    units_df["MUNICÍPIO_norm"] = (
        units_df["MUNICÍPIO"]
        .astype(str)
        .str.strip()
        .str.lower()
        .apply(unidecode)
    )

    units_df["AGUA_ESGOTO_norm"] = (
        units_df["ÁGUA/ESGOTO"]
        .astype(str)
        .str.strip()
        .str.lower()
        .apply(unidecode)
    )

    filtered_units = units_df[units_df["MUNICÍPIO_norm"] == town_name]

    if "esgoto" in inspection_type:
        filtered_units = filtered_units[filtered_units["AGUA_ESGOTO_norm"].str.contains("esgoto")]
    else:
        filtered_units = filtered_units[filtered_units["AGUA_ESGOTO_norm"].str.contains("agua")]

    filtered_units.insert(0, "ITEM", range(1, len(filtered_units) + 1))
    filtered_units["OBSERVAÇÃO"] = ""

    # Seleciona colunas finais
    final_df = filtered_units[["ITEM", "SISTEMA", "UNIDADE", "OBSERVAÇÃO"]]
    table = document.add_table(rows=1, cols=len(final_df.columns))

    hdr_cells = table.rows[0].cells
    for idx, col_name in enumerate(final_df.columns):
        cell = hdr_cells[idx]
        para = cell.paragraphs[0]
        para.text = ""
        try:
            para.style = 'Arial10'
        except Exception:
            pass
        run = para.add_run(str(col_name))
        run.bold = True
        run.font.name = 'Arial'
        run.font.size = Pt(10)
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        para.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        cell._tc.get_or_add_tcPr().append(apply_background_color("D9D9D9"))

    for _, row in final_df.iterrows():
        cells = table.add_row().cells
        for idx, value in enumerate(row):
            cell = cells[idx]
            para = cell.paragraphs[0]
            para.text = ""
            try:
                para.style = 'Arial10'
            except Exception:
                pass
            run = para.add_run(str(value))
            run.font.name = 'Arial'
            run.font.size = Pt(10)

    set_borders_table(table)
    paragraph_index = search_paragraph(document, text)[0]
    paragraph = document.paragraphs[paragraph_index]
    paragraph._element.addnext(table._element)

    print("✅ Tabela de unidades do município criada e inserida.")

def create_documents_table(document, text):
    df_documents = documents_excel.copy()
    table = document.add_table(rows=1, cols=len(df_documents.columns))

    hdr_cells = table.rows[0].cells
    for idx, col_name in enumerate(df_documents.columns):
        para = hdr_cells[idx].paragraphs[0]
        para.text = ""
        run = para.add_run(str(col_name))
        run.bold = True
        run.font.name = "Arial"
        run.font.size = Pt(10)
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        hdr_cells[idx]._tc.get_or_add_tcPr().append(apply_background_color("D9D9D9"))

    # --- Dados ---
    for _, row in df_documents.iterrows():
        cells = table.add_row().cells
        for idx, value in enumerate(row):
            para = cells[idx].paragraphs[0]
            para.text = ""
            valor_texto = "" if pd.isna(value) else str(value)
            run = para.add_run(valor_texto)
            run.font.name = "Arial"
            run.font.size = Pt(10)
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER

    set_borders_table(table)

    list_paragraph_index = search_paragraph(document, text)[0]
    list_paragraph = document.paragraphs[list_paragraph_index]
    list_paragraph._element.addnext(table._element)

    print("✅ Tabela de documentos criada e inserida.")

def create_statistics_table(document):
    df_statistics = town_statistics.copy()
    report_data = get_inspections_data()
    report_town = report_data["Municipio"].upper()
    stats_from_this_town = df_statistics[df_statistics["ANO BASE: 2023"] == report_town]
    selected_columns = [
        "Quantidade de economias residenciais ativas de água (A) - EAA", 
        "Quantidade de economias residenciais inativas de água (B)-EIA", 
        "Quantidade de economias residenciais ativas de esgoto (C) - EAE", 
        "Quantidade de economias residenciais inativas de esgoto (D) - EIE",
        "Quantidade de economias residenciais ativas com tratamento de esgoto (E) - EAT", 
        "Quantidade de economias residenciais inativas com tratamento de esgoto (F) - EIT", 
        "Quantidade de domicílios residenciais existentes na área de abrangência do prestador de serviços (G) - DAP"
    ]

    stats_from_this_town = stats_from_this_town[selected_columns]
    pernambuco_stats = {"Quantidade de economias residenciais ativas de água" : "2.261.695", 
                        "Quantidade de economias residenciais inativas de água" : "377.745",
                        "Quantidade de economias residenciais ativas de esgoto" : "654.143",
                        "Quantidade de economias residenciais inativas de esgoto" : "300.683",
                        "Quantidade de economias residenciais ativas com tratamento de esgoto" : "654.143",
                        "Quantidade de economias residenciais inativas com tratamento de esgoto" : "300.683",
                        "Quantidade de domicílios residenciais existentes na área de abrangência do prestador de serviços" : "2.646.895"}
                        
    table = document.add_table(rows=8, cols=3)
    print(stats_from_this_town)


create_statistics_table()