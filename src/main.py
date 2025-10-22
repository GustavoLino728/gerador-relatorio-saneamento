from report import generate_operational_report, generate_commercial_report
from common.excel import get_inspections_data
from common.utils import sanitize_value

def main():
    report_data = get_inspections_data()
    if report_data is None:
        return None
    inspection_type = sanitize_value(report_data["Tipo da Fiscalização"])

    if inspection_type == "comercial":
        generate_commercial_report()
    else:
        generate_operational_report()

if __name__ == "__main__":
    try:
        main()
    finally:
            input("\nExecução finalizada. Pressione Enter para sair...")