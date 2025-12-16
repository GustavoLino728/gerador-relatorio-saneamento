import os
import sys

def get_base_path():
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

BASE_PATH = get_base_path()

DATA_PATH = os.path.join(BASE_PATH, "data")
REPORTS_PATH = os.path.join(BASE_PATH, "reports")
ASSETS_PATH = os.path.join(BASE_PATH, "assets")

SHEET_PATH = os.path.join(DATA_PATH, "Cadastro das Fiscalizações.xlsm")
