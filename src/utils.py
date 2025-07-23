import os

REPORT_DIR = "../data/"
REPORT_NAME = "RELATÓRIO MODELO"
REPORT_EXT = ".docx"
def next_filename():
    for i in range(1000):
        name = f"{REPORT_NAME}{'' if i == 0 else f' - {i}'}{REPORT_EXT}"
        path = os.path.join(REPORT_DIR, name)
        if not os.path.exists(path):
            return path

def search_paragraph(document ,text):
    paragraphs_found_by_search = []
    for i, p in enumerate(document.paragraphs):
        print(">>> Busca no documento iniciada")
        if text in p.text:
            paragraphs_found_by_search.append(i)
            print(">>> Uma ocorrencia foi encontrada")
    return paragraphs_found_by_search