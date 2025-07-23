from docx import Document
from photos import get_images_from_dir
from photos import divide_photos
from utils import next_filename
from utils import search_paragraph

document = Document(r"../data/RELATÓRIO MODELO.docx")

def generate_report():

    # Busca a Sessão pra inserir as imagens
    search_non_conformity = search_paragraph(document,"APÊNDICE 1 - NÃO CONFORMIDADES")
    last_position = document.paragraphs[search_non_conformity[-1]]

    list_of_images_path = get_images_from_dir()
    divide_photos(document, last_position, list_of_images_path)

    print(">>> Alterações Concluidas")

    new_archive = next_filename() 
    document.save(new_archive)   

