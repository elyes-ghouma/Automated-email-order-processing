
import openpyxl
from functions_real import read_jsonl_file

wb = openpyxl.load_workbook("sage_form_vide_rxl.xlsm",keep_vba=True)

sheet1=wb["Entête Livraison"]
sheet2 = wb["Lignes de Commandes"]
cols1=["D","E","F","K","P","R","S"]
cols2=["D","E","J"]
dc1=13
fc1=14
for i in range(dc1,fc1):
    for col in cols1:
        sheet1[col+str(i)]=""

dc2=14
fc2=31

for i in range(dc2,fc2):
    for col in cols2:
        sheet2[col+str(i)]=""

wb.save("sage_form.xlsm")
