from tables import create_generic_table
from excel import get_inspections_data
import pandas as pd
from paths import SHEET_PATH

spreadsheet = pd.ExcelFile(SHEET_PATH)
services = pd.read_excel(spreadsheet, sheet_name="Atendimentos")

def create_quantity_service_table(document):
    services_data = get_commercial_data()
    
    percentLate = (int(services_data["Atendimentos Fora do Prazo"]))/(int(services_data["Atendimentos Totais"]))*100
    table_2_infos = [["Situação do Atendimento", "Quantidade", "Percentual (%)"],
                    ["No Prazo", services_data[""], 100-percentLate],
                    ["Fora do Prazo", services_data[""], percentLate],
                    ["Total", services_data["Atendimentos Totais"], "100,0%"]]
    
    create_generic_table()