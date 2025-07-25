import pandas as pd

SHEET_PATH="../data/Listagem das NC's - Agua e Esgoto.xlsx"

spreadsheet = pd.ExcelFile(SHEET_PATH)

inspections = pd.read_excel(spreadsheet, sheet_name="Fiscalizações")
non_conformities = pd.read_excel(spreadsheet, sheet_name="Nao-conformidades")
eta_water_nonconformities = pd.read_excel(spreadsheet, sheet_name="NCs ETA")  
eea_water_nonconformities = pd.read_excel(spreadsheet, sheet_name="NCs EEA") 
eea_water_nonconformities = pd.read_excel(spreadsheet, sheet_name="NCs REL e RAP")   
ete_sewage_nonconformities = pd.read_excel(spreadsheet, sheet_name="NCs ETE")  
eee_sewage_nonconformities = pd.read_excel(spreadsheet, sheet_name="NCs EEE")  

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
    print(this_report_id)
    this_report_non_conformities = non_conformities[non_conformities["ID da Fiscalização"] == this_report_id]
    return this_report_non_conformities

print(get_this_report())

# with pd.ExcelWriter(SHEET_PATH, mode="a", if_sheet_exists="replace", engine="openpyxl") as writer:
#     inspections.to_excel(writer, sheet_name="Fiscalizações", index=False)