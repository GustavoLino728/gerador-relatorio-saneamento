import os
from PIL import Image
from docx.shared import Inches, Pt
from docx.oxml.shared import OxmlElement
from docx.oxml.ns import qn
from excel import get_non_conformities 
from utils import set_borders_table, get_images_from_dir, search_paragraph


def build_caption_map(df, col_img="Nome da Foto", col_unit="Unidade", col_desc="Não Conformidade"):
    """
    Constrói um dicionário {nome_foto: legenda} a partir de um dataframe.
    """
    captions = {}
    for _, row in df.iterrows():
        image_name = row[col_img]
        captions[image_name] = f"{image_name} - {row[col_unit]}: {row[col_desc]}"
    return captions


def resize_images(path="../assets", size=(250, 250)):
    """
    Redimensiona a imagem para o tamnho especifico e padrão
    """
    for root, dirs, files in os.walk(path):
        for filename in files:
            if filename.lower().endswith(('.jpg', '.jpeg', '.png')):
                image_path = os.path.join(root, filename)
                img = Image.open(image_path)
                img_resized = img.resize(size, Image.LANCZOS)
                img_resized.save(image_path)


def create_table_images(document, insert_coord, list_of_images_path, captions=None, title_text="Registros Fotográficos"):
    """
    Cria tabela de imagens genérica com legenda.
    - document: Arquivo Word (Objeto)
    - insert_coord: Onde o bloco será inserido
    - list_of_images_path: lista com caminhos das imagens
    - captions: dict {nome_imagem: legenda}, se None usa só o nome
    - title_text: título exibido antes da tabela
    """
    
    space_before = document.add_paragraph()
    space_after = document.add_paragraph()
    space_before_element = space_before._element
    space_after_element = space_after._element

    title = document.add_paragraph(title_text)
    title.style = 'Arial10'

    num_images = len(list_of_images_path)
    num_rows = ((num_images + 1) // 2) * 2   
    images_table = document.add_table(rows=num_rows, cols=2)
    
    for i, img_path in enumerate(list_of_images_path):
        image_line = i // 2 * 2
        subtitle_line = image_line + 1
        column = i % 2
        
        images_table.cell(image_line, column).paragraphs[0].add_run().add_picture(img_path, width=Inches(3.3), height=Inches(2.5))

        image_name = os.path.splitext(os.path.basename(img_path))[0]

        subtitle = captions.get(image_name, f"{image_name} - SEM LEGENDA") if captions else image_name

        cell = images_table.cell(subtitle_line, column)
        paragraph = cell.paragraphs[0]
        paragraph.style = 'Arial10'
        run = paragraph.add_run(subtitle)
        run.font.name = 'Arial'
        run.font.size = Pt(10)

    insert_coord._element.addnext(space_before_element)
    space_before_element.addnext(title._element)
    title._element.addnext(space_after_element)
    space_after_element.addnext(images_table._element)

    set_borders_table(images_table)

    return images_table


def divide_images(document, insert_coord, list_of_images_path, captions=None, block_size=6):
    """Divide imagens em blocos de N (padrão 6) e cria tabelas"""
    for i in range(0, len(list_of_images_path), block_size):
        table_images = list_of_images_path[i:i+block_size]
        insert_coord = create_table_images(document, insert_coord, table_images, captions)
        
        
def create_all_appendix_images(document, text_nc):
    """
    Cria as tabelas de imagens para todos os apêndices, NCs sempre e Condições gerais cria caso haja imagens na pasta (fotos_nao_conformidades).
    - document: objeto Word
    - text_nc: posição no doc onde começam as NCs
    """

    images_by_folder = get_images_from_dir()

    if "fotos_nao_conformidades" in images_by_folder:
        df_nc = get_non_conformities()
        captions_nc = build_caption_map(df_nc)
        divide_images(document, text_nc, images_by_folder["fotos_nao_conformidades"], captions=captions_nc, block_size=6)

    if "fotos_condicoes_gerais" in images_by_folder:    
        text_info = document.paragraphs[search_paragraph(document,"APÊNDICE 2 – CONDIÇÕES GERAIS")[-1]]
        divide_images(document, text_info, images_by_folder["fotos_condicoes_gerais"], captions=None, block_size=6)