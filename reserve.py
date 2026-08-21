import os
import win32com.client as win32

def edit(wb,sheet,col,row,value):
    wb.worksheets(sheet).Cells(col,row).Value=value



# Absolute paths are required by win32com
file_path = os.path.abspath("sage_form.xlsm")

# Launch Excel in background
excel = win32.gencache.EnsureDispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False

try:
    # Open workbook natively
    wb = excel.Workbooks.Open(file_path)

    for sheet in wb.Worksheets:
        try:
            sheet.Unprotect("")
        except Exception:
            print("lule")
            pass

    # Perform updates here if needed (e.g., wb.Worksheets("Sheet1").Cells(2, 1).Value = "Test")
    sheet1="Entête Livraison"
    sheet2 = "Lignes de Commandes"
    cols1=["D","E","F","K","P","R","S"]
    cols2=["D","E","J"]
    dc1=13
    fc1=14
    for i in range(dc1,fc1):
        for col in cols1:
            edit(wb,sheet1,col,i,"")

    dc2=14
    fc2=31

    for i in range(dc2,fc2):
        for col in cols2:
            edit(wb,sheet2,col,i,"")

    # Save and close using native Excel engine
    wb.Save()
    wb.Close()
finally:
    excel.Quit()