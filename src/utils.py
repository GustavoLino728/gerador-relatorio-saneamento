import os
from excel import get_inspections_data

REPORT_DIR = "../reports/"
REPORT_NAME = "RELATÓRIO MODELO"
REPORT_EXT = ".docx"
def next_filename():
    for i in range(1, 1000):
        name = f"{REPORT_NAME}{'' if i == 0 else f' - {i}'}{REPORT_EXT}"
        path = os.path.join(REPORT_DIR, name)
        if not os.path.exists(path):
            return path

def search_paragraph(document, text):
    paragraphs_found_by_search = []
    for i, p in enumerate(document.paragraphs):
        print(">>> Busca no documento iniciada")
        if text in p.text:
            paragraphs_found_by_search.append(i)
            print(">>> Uma ocorrencia foi encontrada")
    return paragraphs_found_by_search

def substitute_placeholders(document):
    excel_data = get_inspections_data()
    for p in document.paragraphs:
        for key, value in excel_data.items():
            placeholder = f"{{{{{key}}}}}"
            if placeholder in p.text:
                p.text = p.text.replace(placeholder, str(value))

    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for key, value in excel_data.items():
                        placeholder = f"{{{{{key}}}}}"
                        if placeholder in p.text:
                            p.text = p.text.replace(placeholder, str(value))
