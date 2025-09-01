import os
from PIL import Image
from docx.shared import Inches, Pt
from docx.oxml.shared import OxmlElement
from docx.oxml.ns import qn
from excel import get_non_conformities 
from utils import set_borders_table


def get_images_from_dir(path="./assets"):
    """Retorna o caminho de todas as fotos que estão na pasta assets (.jpg, .jpeg, .png)"""
    all_files = os.listdir(path)
    all_images = [f for f in all_files if f.lower().endswith(('.jpg', '.jpeg', '.png'))]
    return [os.path.join(path, f) for f in all_images]


# def resize_images(path="../assets", size=(500, 500)):
#     for filename in os.listdir(path):
#         if filename.lower().endswith(('.jpg', '.jpeg', '.png')):
#             image_path = os.path.join(path, filename)
#             img = Image.open(image_path)
#             img_resized = img.resize(size, Image.LANCZOS)
#             img_resized.save(image_path)
#     print(">>> Imagens redimensionadas para 210x210")


def create_table_images(document, last_position, list_of_images_path, df_non_conformities):
    """
    Cria a tabela de fotos das não conformidades, com legenda e subtitulo
    - document: Arquivo Word (Objeto)
    - last_position: A posição onde começa a ser inserido as imagens
    - list_of_images_path: Lista dos caminhos de todas as imagens
    - df_non_conformities: Dataframe das não conformidades usadas para a legenda das imagens
    """

    space_before = document.add_paragraph()
    space_after = document.add_paragraph()
    space_before_element = space_before._element
    space_after_element = space_after._element

    title = document.add_paragraph("Registros Fotográficos das Não Conformidades")
    title.style = 'Arial10'
    images_table = document.add_table(rows=6, cols=2)
    
    for i in range(len(list_of_images_path)):
        image_line = i // 2 * 2
        subtitle_line = image_line + 1
        collum = i % 2

        images_table.cell(image_line, collum).paragraphs[0].add_run().add_picture(list_of_images_path[i], width=Inches(4.3))
        image_name = os.path.splitext(os.path.basename(list_of_images_path[i]))[0]
        line = df_non_conformities[df_non_conformities["Nome da Foto"] == image_name]

        if not line.empty:
            unit = line.iloc[0]["Unidade"]
            description = line.iloc[0]["Não Conformidade"]
            subtitle = f"{image_name} - {unit}: {description}"
        else:
            subtitle = f"{image_name} - NÃO ENCONTRADO"

        cell = images_table.cell(subtitle_line, collum)
        paragraph = cell.paragraphs[0]
        paragraph.style = 'Arial10'
        run = paragraph.add_run(subtitle)
        run.font.name = 'Arial'
        run.font.size = Pt(10)

    last_position._element.addnext(space_before_element)
    space_before_element.addnext(title._element)
    title._element.addnext(space_after_element)
    space_after_element.addnext(images_table._element)

    set_borders_table(images_table)
    print(">>> Tabela de Fotos Criada")

    return images_table


def divide_images(document, last_position, list_of_images_path):
    """Divide as imagens em blocos de 6 imagens cada"""
    for i in range(0, len(list_of_images_path), 6):
        table_images = list_of_images_path[i:i+6]
        last_position = create_table_images(document, last_position, table_images, get_non_conformities())
        print(f">>> {i}° bloco concluido")
