import openpyxl
from functions_real import read_jsonl_file, read_json_file, day_name, stocker_details, add_to_order

def remplir_excel(order_id):
    wb = openpyxl.load_workbook("sage_form_vide_rxl.xlsm",keep_vba=True)

    param=read_json_file("parametres.json")
    path=param["order_directory_path"]+"Order"+str(order_id)+"/"

    date_data=read_json_file(path+"date"+str(order_id)+".json")

    detail_data=read_jsonl_file(path+"detail"+str(order_id)+".jsonl")

    order_data=read_json_file(path+"order"+str(order_id)+".json")
    print(order_data)
    stocker_data=stocker_details("SAGEMCOM",order_data["stockeur"])
    order_data=add_to_order(order_data,stocker_data)

    sheet1=wb["Entête Livraison"]
    cols1=["D","E","F","K","P","R","S"]
    c1="13"
    data1=[order_data["num_stockeur"],
        date_data["no_vr"],
        day_name(date_data["date"]),
        order_data["num_stockeur"],
        order_data["num_dep"],
        order_data["ville"],
        order_data["pays"]
    ]
    for i in range(7):
        sheet1[cols1[i]+c1]=data1[i]

    sheet2 = wb["Lignes de Commandes"]
    cols2=["D","E","J"]
    c2=14
    for line in detail_data:
        data2=[date_data["no_vr"],
            line["Réf SAGEM"],
            line["Qté à commander (en UC)"]
        ]
        for i in range(3):
            sheet2[cols2[i]+str(c2)]=data2[i]
        c2+=1

    wb.save(path+"sage_form"+str(order_id)+".xlsm")

import ollama
response=ollama.list()
for model in response['models']:
    name = model['model']
    print(name)