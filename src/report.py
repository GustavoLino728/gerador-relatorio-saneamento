from common.images import create_all_appendix_images
from common.utils import substitute_placeholders, next_filename, search_paragraph, decide_report_type, insert_general_condition_section, is_file_open
# from common.excel import mark_report_as_finished, get_commercial_data
from common.tables import create_non_conformities_table, create_town_units_table, create_documents_table, create_general_information_table, create_abbreviations_table, create_last_report_table
from operational.tables import create_statistics_table, create_quality_index_table, create_table_7
from commercial.tables import create_quantity_service_table, create_late_service_reason_table
from commercial.analysis import analyze_deadline_with_reason
from common.paths import SHEET_PATH
from tqdm import tqdm


def generate_operational_report():
    """Função responsável por gerar os relátorios de fiscalizações do tipo operacional (Agua e Esgoto)"""

    document = decide_report_type()
    if document is None:
        print("❌ Nenhum relatório pendente para gerar.")
        return

    is_file_open(SHEET_PATH)
    
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
        ("Inserir seção de Condições gerais", lambda: insert_general_condition_section(document, "APÊNDICE 1 - NÃO CONFORMIDADES")),
        ("Inserir imagens", lambda: create_all_appendix_images(document, document.paragraphs[search_paragraph(document,"APÊNDICE 1 - NÃO CONFORMIDADES")[-1]])),
        ("Substituir placeholders", lambda: substitute_placeholders(document)),
        ("Salvar documento", lambda: document.save(next_filename())),
        # ("Marcando Como Finalizado", lambda: mark_report_as_finished()),
    ]

    for desc, func in tqdm(steps, desc="Gerando relatório", unit="etapa"):
        func()

    print("✅ Relátorio Gerado com Sucesso!")
    
def generate_commercial_report():
    """Função responsável por gerar os relátorios de fiscalizações do tipo comercial"""
    document = decide_report_type()
    analysis_result = analyze_deadline_with_reason()
    
    if document is None:
        print("❌ Nenhum relatório pendente para gerar.")
        return

    is_file_open(SHEET_PATH)
    
    steps = [
        ("Siglas e Abreviações", lambda: create_abbreviations_table(document, "LISTA DE ABREVIATURAS E SIGLAS")),
        ("Informações gerais", lambda: create_general_information_table(document, "3.	INFORMAÇÕES GERAIS")),
        ("Não conformidades", lambda: create_non_conformities_table(document, "Tabela 1 - Lista de NCs da Loja de atendimento {{Municipio}}.")),
        ("Quantidade de Atendimentos", lambda: create_quantity_service_table(document, analysis_result)),
        ("Motivo de Encerramento", lambda: create_late_service_reason_table(document, analysis_result)),
        ("Inserir imagens", lambda: create_all_appendix_images(document, document.paragraphs[search_paragraph(document,"APÊNDICE 1 - NÃO CONFORMIDADES")[-1]])),
        ("Substituir placeholders", lambda: substitute_placeholders(document)),
        ("Substituir placeholders (Especificos de Comercial)", lambda: substitute_placeholders(document, excel_data=analysis_result)),
        ("Salvar documento", lambda: document.save(next_filename())),
        # ("Marcando Como Finalizado", lambda: mark_report_as_finished())
    ]

    for desc, func in tqdm(steps, desc="Gerando relatório", unit="etapa"):
        func()

    print("✅ Relátorio Gerado com Sucesso!")