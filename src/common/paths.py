import os


BASE_PATH = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

DATA_PATH = os.path.join(BASE_PATH, "data")
REPORTS_PATH = os.path.join(BASE_PATH, "reports")
ASSETS_PATH = os.path.join(BASE_PATH, "assets")

SHEET_PATH = os.path.join(DATA_PATH, "Cadastro das Fiscalizações.xlsm")