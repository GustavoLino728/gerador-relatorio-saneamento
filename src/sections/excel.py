import pandas as pd
from unidecode import unidecode

SHEET_PATH="../data/Listagem das NC's - Agua e Esgoto.xlsm"
spreadsheet = pd.ExcelFile(SHEET_PATH)

inspections = pd.read_excel(spreadsheet, sheet_name="Fiscalizações")
non_conformities = pd.read_excel(spreadsheet, sheet_name="Nao-conformidades")
list_of_all_units = pd.read_excel(spreadsheet, sheet_name="Lista-SES-e-SAA")
documents_excel = pd.read_excel(spreadsheet, sheet_name="Envio de Documentos")
town_statistics = pd.read_excel(spreadsheet, sheet_name="Estatisticas ")

ete_sewage_nonconformities = pd.read_excel(spreadsheet, sheet_name="NCs ETE")
eee_sewage_nonconformities = pd.read_excel(spreadsheet, sheet_name="NCs EEE")
eta_water_nonconformities = pd.read_excel(spreadsheet, sheet_name="NCs ETA")
eea_water_nonconformities = pd.read_excel(spreadsheet, sheet_name="NCs REL e RAP")

def get_this_report():
    not_done_reports = inspections[inspections["Relatório Gerado"] == "Gerar"]
    if not not_done_reports.empty:
        return not_done_reports.index[0]
    else:
        raise ValueError("Todos os relatórios já foram gerados.")

def get_inspections_data():
    this_report = get_this_report()
    if this_report >= len(inspections):
        raise IndexError("O índice calculado está fora do alcance do DataFrame de inspeções.")
    return inspections.iloc[this_report].to_dict()

def get_non_conformities():
    this_report_id = get_this_report()
    this_report_non_conformities = non_conformities[non_conformities["ID da Fiscalização"] == this_report_id]
    this_report_non_conformities["Sigla"] = this_report_non_conformities["Unidade"].str.extract(r'^(...)\s*-')
    return this_report_non_conformities

print(get_inspections_data())

# def all_unities_town_table():
#     data = get_inspections_data()
#     list_of_all_units.columns = list_of_all_units.columns.str.strip()
#     tipo_fiscalizacao = unidecode(str(data["Tipo da Fiscalização"]).strip().lower())
#     town = str(data["Municipio"]).upper()

#     if tipo_fiscalizacao == "agua":
#         list_of_saa = list_of_all_units[
#             (list_of_all_units["MUNICÍPIO"] == town) & (list_of_all_units["ÁGUA/ESGOTO"] == "agua")
#         ]
#         print(list_of_saa)

#     elif tipo_fiscalizacao == "esgoto":
#         list_of_ses = list_of_all_units[
#             (list_of_all_units["MUNICÍPIO"] == town) & (list_of_all_units["ÁGUA/ESGOTO"] == "esgoto")
#         ]
#         print(list_of_ses)
    
# all_unities_town_table()
    

# with pd.ExcelWriter(SHEET_PATH, mode="a", if_sheet_exists="replace", engine="openpyxl") as writer:
#     inspections.to_excel(writer, sheet_name="Fiscalizações", index=False)