import pandas as pd
from unidecode import unidecode
from utils import sanitize_value


SHEET_PATH="./data/Listagem das NC's - Agua e Esgoto.xlsm"
spreadsheet = pd.ExcelFile(SHEET_PATH)

inspections = pd.read_excel(spreadsheet, sheet_name="Fiscalizações")
non_conformities = pd.read_excel(spreadsheet, sheet_name="Nao-conformidades", header=5)
list_of_all_units = pd.read_excel(spreadsheet, sheet_name="Lista-SES-e-SAA")
documents_excel = pd.read_excel(spreadsheet, sheet_name="Envio de Documentos")
town_statistics = pd.read_excel(spreadsheet, sheet_name="Estatisticas ")

ete_sewage_nonconformities = pd.read_excel(spreadsheet, sheet_name="NCs ETE")
eee_sewage_nonconformities = pd.read_excel(spreadsheet, sheet_name="NCs EEE")
eta_water_nonconformities = pd.read_excel(spreadsheet, sheet_name="NCs ETA")
eea_water_nonconformities = pd.read_excel(spreadsheet, sheet_name="NCs REL e RAP")

def get_this_report():
    """Retorna o id referente a fiscalização atual baseado em qual linha estiver escrito [Gerar]"""
    reports = inspections["Relatório Gerado"].apply(sanitize_value)
    not_done_reports = inspections[reports == "gerar"]
    if not not_done_reports.empty:
        return not_done_reports.index[0]
    else:
        raise ValueError("Todos os relatórios já foram gerados.")


def get_inspections_data():
    """Retorna os dados da fiscalização atual"""
    this_report = get_this_report()
    if this_report >= len(inspections):
        raise IndexError("O índice calculado está fora do alcance do DataFrame de inspeções.")
    return inspections.iloc[this_report].to_dict()


def get_non_conformities():
    """Retorna as não conformidades do relatório atual que vai ser gerado"""
    this_report_id = get_this_report()
    this_report_non_conformities = non_conformities[non_conformities["ID da Fiscalização"] == this_report_id]
    this_report_non_conformities["Sigla"] = this_report_non_conformities["Unidade"].str.extract(r'^(.*?)\s*-')
    return this_report_non_conformities
    
# with pd.ExcelWriter(SHEET_PATH, mode="a", if_sheet_exists="replace", engine="openpyxl") as writer:
#     inspections.to_excel(writer, sheet_name="Fiscalizações", index=False)