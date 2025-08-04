from docx import Document
from images import get_images_from_dir
from images import divide_images
from utils import substitute_placeholders
# from images import resize_images
from utils import next_filename
from utils import search_paragraph

document = Document(r"../data/RELATÓRIO MODELO.docx")

def generate_report():
    # Busca a Sessão pra inserir as imagens
    search_non_conformity = search_paragraph(document,"APÊNDICE 1 - NÃO CONFORMIDADES")
    last_position = document.paragraphs[search_non_conformity[-1]]

    # resize_images()

    list_of_images_path = get_images_from_dir()
    divide_images(document, last_position, list_of_images_path)
    substitute_placeholders(document)
    
    print(">>> Alterações Concluidas")

    new_archive = next_filename()
    document.save(new_archive) 