from common.tables import create_generic_table
from common.paths import SHEET_PATH
import pandas as pd

spreadsheet = pd.ExcelFile(SHEET_PATH)

def create_quantity_service_table(document, analysis_result, text="Tabela 2 - Quantidade de atendimentos"):
    """Cria a tabela 2, descrevendo quantos atendimentos houveram naquela loja e quantos (%) foram fora do prazo e no prazo"""
    table_2_info = [["Situação do Atendimento", "Quantidade", "Percentual (%)"],
                    ["No Prazo", analysis_result["Quantidade dentro do prazo"], analysis_result["% dentro do prazo"]],
                    ["Fora do Prazo", analysis_result["Quantidade fora do prazo"], analysis_result["% fora do prazo"]],
                    ["Total", analysis_result["Quantidade total de atendimentos"], "100,0%"]]
    
    create_generic_table(document, rows_data=table_2_info, text_after_paragraph=text, col_widths=[4, 2, 2])


def create_late_service_reason_table(document, analysis_result, text="Tabela 3 - Motivo do encerramento"):
    """"""
    reasons_counts = analysis_result["Contagem Motivos Fora do Prazo"]
    sorted_reasons = sorted(reasons_counts.items(), key=lambda x: x[1], reverse=True)

    top_5 = dict(sorted_reasons[:5])
    others = sorted_reasons[5:]
    others_sum = sum(v for _, v in others)

    if others_sum > 0:
        top_5["OUTROS"] = others_sum

    table_data = [["Motivo do Encerramento", "Quantidade"]]
    for motivo, count in top_5.items():
        table_data.append([motivo, int(count)])

    create_generic_table(document, rows_data=table_data, text_after_paragraph=text, col_widths=[4, 2])
    
    if others:
        document.add_paragraph("Detalhamento dos motivos incluídos em 'OUTROS':")
        for motivo, count in others:
            document.add_paragraph(f"- {motivo}: {count}")