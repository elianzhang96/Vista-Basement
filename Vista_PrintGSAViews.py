from openpyxl import load_workbook
from openpyxl.styles import Alignment
from gsapy import GSA
from datetime import datetime
import time
import Lib_PrintGSAViews as lib
import os
from openpyxl.utils import get_column_letter
import matplotlib.pyplot as plt
import subprocess


# GSA model and outputs - Building model
gsa_file_path = r'C:\Users\Elian.Zhang\Arup\PROJECT VISTA - 4 Internal Data\04 Calculations\07 Structural\09 Substructure\Stage 4\12 Basement Design\01_GSA\03_SIMPLIFIED' 
gsa_file_name = 'Semplified slab design  B1_modified_v10.gwb'

views = ['RC Design - Wind Leading - Top A reinf.', 'RC Design - Wind Leading - Top B reinf.', 'RC Design - Wind Leading - Bot. A reinf.', 'RC Design - Wind Leading - Bot. B reinf.', 'RC Design - Live Leading - Top A reinf.', 'RC Design - Live Leading - Top B reinf.', 'RC Design - Live Leading - Bot. A reinf.', 'RC Design - Live Leading - Bot. B reinf.', 'RC Design - Shrinkage - Top A reinf.', 'RC Design - Shrinkage - Top B reinf.', 'RC Design - Shrinkage - Bot. A reinf.', 'RC Design - Shrinkage - Bot. B reinf.',
]


##################### SCRIPT #####################


##################### GET RESULTS FROM GSA MODEL

print(datetime.now().replace(microsecond=0), ": start.")


gsa_file = gsa_file_path + "\\" + gsa_file_name
model = GSA(gsa_file, version="10.2")

print(datetime.now().replace(microsecond=0), ": model opened.")

for i in views:
    model.save_view_to_file(i, file_type='PNG')
    print(datetime.now().replace(microsecond=0), ": ", i, " printed.")

print(datetime.now().replace(microsecond=0), ": all views printed.")


################# OPEN FOLDER

output_folder = gsa_file_path
subprocess.Popen(['explorer', output_folder], shell=True)
















    



