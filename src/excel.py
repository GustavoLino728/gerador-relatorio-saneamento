import pandas as pd
from unidecode import unidecode


SHEET_PATH="./data/Listagem das NC's - Agua e Esgoto.xlsm"
spreadsheet = pd.ExcelFile(SHEET_PATH)

inspections = pd.read_excel(spreadsheet, sheet_name="Fiscalizações")
non_conformities = pd.read_excel(spreadsheet, sheet_name="Nao-conformidades", header=5)
documents_excel = pd.read_excel(spreadsheet, sheet_name="Envio de Documentos")
town_statistics = pd.read_excel(spreadsheet, sheet_name="Estatisticas ")
units_df = pd.read_excel(spreadsheet, sheet_name="Cadastrar Unidades", header=3)

ete_sewage_nonconformities = pd.read_excel(spreadsheet, sheet_name="NCs ETE")
eee_sewage_nonconformities = pd.read_excel(spreadsheet, sheet_name="NCs EEE")
eta_water_nonconformities = pd.read_excel(spreadsheet, sheet_name="NCs ETA")
eea_water_nonconformities = pd.read_excel(spreadsheet, sheet_name="NCs REL e RAP")

def get_this_report():
    """Retorna o id referente a fiscalização atual baseado em qual linha estiver escrito [Gerar]"""
    not_done_reports = inspections[
        inspections["Relatório Gerado"].str.lower() == "gerar"
    ]
    if not not_done_reports.empty:
        return not_done_reports.index[0]
    else:
        print("❌ Todos os relatórios já foram gerados.")


def get_inspections_data():
    """Retorna os dados da fiscalização atual"""
    this_report = get_this_report()
    if this_report >= len(inspections):
        print("❌ As informações da fiscalização não foram encontradas, verifique a aba (Fiscalizações) na planilha e verifique o (ID da fiscalização) e se Relátorio Gerado contém um (Gerar)")
    data = inspections.iloc[this_report].to_dict()
    
    if "Tipo da Fiscalização" in data:
        report_type = str(data["Tipo da Fiscalização"]).strip().lower()
        report_type = unidecode(report_type)       
        report_type = report_type.replace(" ", "") 
        data["Tipo da Fiscalização"] = report_type
        
    if data["Tipo da Fiscalização"] == "agua":
        data["SAA ou SEE"] = "SAA"
    elif data["Tipo da Fiscalização"] == "esgoto":
        data["SAA ou SEE"] = "SEE"
    else:
        print("❌ Tipo de Fiscalização não válido, insira um válido: Agua ou Esgoto")
    return data


def get_non_conformities():
    """Retorna as não conformidades do relatório atual que vai ser gerado"""
    this_report_id = get_this_report()
    this_report_non_conformities = non_conformities[non_conformities["ID da Fiscalização"] == this_report_id].copy()
    this_report_non_conformities["Sigla"] = this_report_non_conformities["Unidade"].str.extract(r'^(.*?)\s*-')
    return this_report_non_conformities
    
# with pd.ExcelWriter(SHEET_PATH, mode="a", if_sheet_exists="replace", engine="openpyxl") as writer:
#     inspections.to_excel(writer, sheet_name="Fiscalizações", index=False)