from common.excel import get_inspections_data, get_non_conformities, town_statistics
from common.utils import sanitize_value, insert_table_7_text
from common.tables import create_generic_table


def create_statistics_table(document, text):
    """
    Cria a tabela 4 - de Informações sobre Pernambuco em Geral (Fixa) e o Municipio da fiscalização (que é retirado da planilha - Estatisticas). 
    Puxa os dados sobre EIA, EAE, EIE, EAT, EIT, DAP
    Formato: 8 linhas x 3 colunas.
    Linha 1: cabeçalho (INFORMAÇÃO, PERNAMBUCO, "Município")
    Linha 2-8: valores correspondentes
    """
    df_statistics = town_statistics.copy()
    report_data = get_inspections_data()
    report_town = report_data["Municipio"]
    report_town = report_town = sanitize_value(report_town)
    report_town = report_town.upper()

    pernambuco_stats = {
        "Quantidade de economias residenciais ativas de água (A) - EAA": "2.261.695",
        "Quantidade de economias residenciais inativas de água (B)-EIA": "377.745",
        "Quantidade de economias residenciais ativas de esgoto (C) - EAE": "654.143",
        "Quantidade de economias residenciais inativas de esgoto (D) - EIE": "300.683",
        "Quantidade de economias residenciais ativas com tratamento de esgoto (E) - EAT": "654.143",
        "Quantidade de economias residenciais inativas com tratamento de esgoto (F) - EIT": "300.683",
        "Quantidade de domicílios residenciais existentes na área de abrangência do prestador de serviços (G) - DAP": "2.646.895"
    }

    columns = list(pernambuco_stats.keys())
    town_stats = df_statistics[df_statistics["ANO BASE: 2023"] == report_town]
    town_stats = town_stats[columns]

    rows_data = [["INFORMAÇÃO", "PERNAMBUCO", report_town]]
    for column in columns:
        town_value = ""
        if not town_stats.empty:
            town_value = town_stats.iloc[0][column]
            if isinstance(town_value, (int, float)):
                town_value = f"{int(town_value):,}".replace(",", ".")
        rows_data.append([column, pernambuco_stats[column], town_value])

    create_generic_table(document=document, rows_data=rows_data, text_after_paragraph=text, col_widths=[6, 1.5, 1.5], align_left=True, font_size=10)


def create_quality_index_table(document, text):
    """
    Cria a tabela 5 - de indicadores de qualidade para o município.
    Formato: 2 linhas x 8 colunas.
    Linha 1: cabeçalho ("Município" + indicadores)
    Linha 2: valores correspondentes
    """
    df_statistics = town_statistics.copy()
    report_data = get_inspections_data()
    report_town = report_data["Municipio"]
    report_town = sanitize_value(report_town)
    report_town = report_town.upper()

    df_statistics.columns = df_statistics.columns.str.replace(r"\s+\(\%\)", "(%)", regex=True)
    df_statistics.columns = df_statistics.columns.str.strip()

    selected_columns = ["IUA(%)", "IUE(%)", "IUT(%)", "ICA", "ICE", "IPD", "IQAP"]

    stats_from_this_town = df_statistics[df_statistics["ANO BASE: 2023"] == report_town]

    valores_tabela = []
    for col in selected_columns:
        valor = "-"
        if not stats_from_this_town.empty and col in stats_from_this_town.columns:
            valor = stats_from_this_town.iloc[0][col]
            if isinstance(valor, (int, float)):
                valor = str(int(valor))
        valores_tabela.append(valor)

    rows_data = [["Município"] + selected_columns, [report_town] + valores_tabela]

    create_generic_table(document, rows_data, text, col_widths=[3] + [1]*7, align_left=True, font_size=10)
    
def create_water_params_table(document, text):
    """
    Cria a tabela de Parametros de Agua, referente a tabela 7 do relatorio de agua. Aparece caso haja pelo menos uma unidade ETA entre as NCs. 
    Lista essas unidades e o analista preenche os campos (CLORO (mg.L ), TURBIDEZ (NTU), OBSERVAÇÕES)
    Formato: 1+N linhas x 4 colunas.
    Linha 1: cabeçalho (QUALIDADE DA ÁGUA (UNIDADES), CLORO (mg.L ), TURBIDEZ (NTU), OBSERVAÇÕES)
    Linha 2-N: Dados correspondentes
    """
    df_ncs = get_non_conformities()
    df_eta = df_ncs[df_ncs["Sigla"] == "ETA"]

    if df_eta.empty:
        return
    
    columns = ["QUALIDADE DA ÁGUA (UNIDADES)", "CLORO (mg.L )", "TURBIDEZ (NTU)", "OBSERVAÇÕES"]
    rows_data = [columns]

    for _, row in df_eta.iterrows():
        row_list = [
            row["Unidade"],  
            "",              
            "",             
            ""               
        ]
        rows_data.append(row_list)

    insert_table_7_text(document)
    create_generic_table(document=document, rows_data=rows_data, text_after_paragraph=text, col_widths=[5, 1.5, 1.5, 3], align_left=False)
    
def create_sewage_params_table(document, text):
    """
    Cria a tabela de Parametros de Qualidade do Efluente. Referente a tabela 7 do relatorio de esgoto. Aparece caso haja pelo menos uma unidade ETE entre as NCs. Lista essas unidades e o analista preenche os campos 
    Formato: 1+N linhas x 3 colunas.
    Linha 1: cabeçalho (QUALIDADE DO EFLUENTE (UNIDADES), DBO filtrada(mg O2/L), OBSERVAÇÕES)
    Linha 2-N: Dados correspondentes
    """
    df_ncs = get_non_conformities()
    df_ete = df_ncs[df_ncs["Sigla"] == "ETE"]

    if df_ete.empty:
        return
    
    columns = ["QUALIDADE DO EFLUENTE (UNIDADES)", "DBO filtrada (mg O2/L)", "OBSERVAÇÕES"]
    rows_data = [columns]

    for _, row in df_ete.iterrows():
        row_list = [
            row["Unidade"],  
            "",              
            ""               
        ]
        rows_data.append(row_list)

    insert_table_7_text(document)
    create_generic_table(document=document, rows_data=rows_data, text_after_paragraph=text, col_widths=[5, 2, 3], align_left=False)

def create_table_7(document):
    """Decide qual das tabelas 7 deve ser gerada com base no tipo da fiscalização"""
    report_data = get_inspections_data()
    inspection_type = sanitize_value(report_data["Tipo da Fiscalização"])
    if inspection_type == 'agua':
        create_water_params_table(document, "Tabela 7 - Parâmetros da qualidade da água.")
    elif inspection_type == 'esgoto':
        create_sewage_params_table(document, "Tabela 7 - Parâmetros da qualidade do efluente.")