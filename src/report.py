from images import get_images_from_dir
from images import divide_images
from utils import substitute_placeholders, next_filename, search_paragraph, decide_report_type
from tables import create_non_conformities_table, create_town_units_table, create_documents_table, create_statistics_table, create_quality_index_table, create_general_information_table, create_abbreviations_table, create_last_report_table, create_table_7
from tqdm import tqdm
# from images import resize_images

def generate_report():
    """Função principal, gera todas as tabelas e substitui os placeholders"""

    document = decide_report_type()

    steps = [
        ("Siglas e Abreviações", lambda: create_abbreviations_table(document, "LISTA DE ABREVIATURAS E SIGLAS")),
        ("Informações gerais", lambda: create_general_information_table(document, "3.	INFORMAÇÕES GERAIS")),
        ("Documentos", lambda: create_documents_table(document, "Tabela 1 - Principais documentações solicitadas.")),
        ("Unidades do município", lambda: create_town_units_table(document, "Tabela 2 - Descrição dos {{SAA ou SEE}} {{Municipio}}.")),
        ("Último relatório", lambda: create_last_report_table(document, "Tabela 3 - Contexto histórico resumido das fiscalizações do município de {{Municipio}}.")),
        ("Estatísticas", lambda: create_statistics_table(document, "Tabela 4 - Informações do prestador de serviços e do município de {{Municipio}}.")),
        ("Índices de qualidade", lambda: create_quality_index_table(document, "Tabela 5 - Principais Indicadores Regulatórios do município {{Municipio}}.")),
        ("Não conformidades", lambda: create_non_conformities_table(document, "Tabela 6 - Lista de NCs do {{SAA ou SEE}} {{Municipio}}")),
        ("Tabela 7", lambda: create_table_7(document)),
        ("Inserir imagens", lambda: divide_images(document, document.paragraphs[search_paragraph(document,"APÊNDICE 1 - NÃO CONFORMIDADES")[-1]], get_images_from_dir())),
        ("Substituir placeholders", lambda: substitute_placeholders(document)),
        ("Salvar documento", lambda: document.save(next_filename())),
    ]

    for desc, func in tqdm(steps, desc="Gerando relatório", unit="etapa"):
        func()

    print("✅ Relátorio Gerado com Sucesso!")