import os
import sys


BASE_PATH = os.path.dirname(os.path.abspath(sys.argv[0]))

DATA_PATH = os.path.join(BASE_PATH, "data")       
REPORTS_PATH = os.path.join(BASE_PATH, "reports") 
ASSETS_PATH = os.path.join(BASE_PATH, "assets")   

SHEET_PATH= os.path.join(DATA_PATH, "Listagem das NC's - Agua e Esgoto.xlsm")