from langchain_core.prompts import ChatPromptTemplate
from langchain_ollama.llms import OllamaLLM
from email.message import EmailMessage
from dotenv import load_dotenv
from datetime import datetime
from ast import literal_eval
from flask import send_file
from dateutil import parser
from time import sleep
from copy import copy
import pandas as pd
import openpyxl
import imaplib
import smtplib
import ollama
import email
import json
import csv
import os
import re


#--------------------------Initialazing functions-----------------------------


def init_client(path):
    file_path=path+"/Clients.jsonl"
    if not os.path.isfile(file_path):
        empty_file(file_path)


def set_main_folder_path(path):
    parametres=read_json_file("parametres.json")
    
    parametres["main_directory_path"]=create_directory(path,"Vectorys Planification")

    parametres["orders_directory_path"]=create_directory(parametres["main_directory_path"],"Orders")

    parametres["excel_files_directory_path"]=create_directory(parametres["main_directory_path"],"Excel Files")

    create_json_file(parametres,"parametres.json")


#---------------------------files treatment functions---------------------------


def file_existance(file_name):
    return os.path.isfile(file_name)


def read_json_file(file_name):
    """Safely reads a JSON file using a context manager.

    Returns an empty dict if the file is missing or empty.
    """
    if not os.path.exists(file_name) or os.path.getsize(file_name) == 0:
        return {}
    try:
        with open(file_name, "r", encoding="utf-8",errors="ignore") as f:
            return json.load(f)
    except json.JSONDecodeError:
        return {}


def read_jsonl_file(file_name):
    """Reads a JSON Lines (.jsonl) file and parses each line into a dict."""
    if not os.path.exists(file_name):
        return []

    data = []
    with open(file_name, "r", encoding="utf-8",errors="ignore") as f:
        for line in f:
            line = line.strip()
            if line:  # Skip blank lines
                try:
                    data.append(json.loads(line))
                except json.JSONDecodeError:
                    continue
    return data


def dict_to_json(dict_data):
    """Converts a dict to a formatted JSON string."""
    return json.dumps(dict_data, ensure_ascii=False)


def json_to_dict(json_data):
    """Safely parses a JSON string into a dictionary."""
    if not json_data or not isinstance(json_data, str):
        return {}

    try:
        parsed = json.loads(json_data.strip(), strict=False)
        return parsed if isinstance(parsed, dict) else {"data": parsed}
    except json.JSONDecodeError:
        return {}


def append_jsonl_file(dict_data,jsonl_file_name):
    jsonl_file=open(jsonl_file_name,"a",encoding="utf-8",errors="ignore")
    jsonl_file.write(json.dumps(dict_data, ensure_ascii=False)+"\n")
    jsonl_file.close()


def create_jsonl_file(list_dict_data,jsonl_file_name):
    empty_file(jsonl_file_name)
    for dict_data in list_dict_data:
        append_jsonl_file(dict_data,jsonl_file_name)


def create_json_file(dict_data,json_file_name):
    json_file=open(json_file_name,"w",encoding="utf-8",errors="ignore")
    json.dump(dict_data, json_file, ensure_ascii=False, default=str)
    json_file.close()


def append_json_to_csv(dict_json_data,csv_file_name):
    if dict_json_data:
        if type(dict_json_data)!=dict:
            data=json_to_dict(dict_json_data)
        else:
            data=dict_json_data

        row_data = {}
        '''
        nb_items=len(data.get("items",[]))
        i=0
        if nb_items==0:'''
        for key, value in data.items():
            if isinstance(value, (dict, list)):
                row_data[key] = json.dumps(value, ensure_ascii=False) 
            else:
                row_data[key] = value

        file_exist=os.path.isfile(csv_file_name)
        csv_file=open(csv_file_name,"a",newline="",encoding="utf-8-sig")
        fieldnames = list(row_data.keys())
        writer=csv.DictWriter(csv_file,fieldnames=fieldnames)
        if not file_exist:
            writer.writeheader()
        writer.writerow(row_data)
        csv_file.close()
        '''
        else:
            for item in data["items"] :
                i+=1
                for key, value in data.items():
                    #if isinstance(value, (dict, list)):
                    if key=="items":
                        for key_item,value_item in item.items():
                            row_data[key_item]=value_item
                        #row_data[key] = json.dumps(value, ensure_ascii=False) 10
                    else:
                        row_data[key] = value

                file_exist=os.path.isfile(csv_file_name)
                csv_file=open(csv_file_name,"a",newline="",encoding="utf-8-sig")
                fieldnames = list(row_data.keys())
                writer=csv.DictWriter(csv_file,fieldnames=fieldnames)
                if not file_exist:
                    writer.writeheader()
                writer.writerow(row_data)
                csv_file.close()'''


def xl_to_csv(xl_file,csv_file):
    # Load the Excel file
    
    if xl_file.find(".xlsx")!=-1 or xl_file.find(".xlsm")!=-1:
        # For .xlsx or .xlsm files:
        try:
            df = pd.read_excel(xl_file, engine="openpyxl")  
        except:
            df = pd.read_excel(xl_file, engine="xlrd")
    elif xl_file.find(".xls")!=-1:
        # For .xls (older Excel format):
        df = pd.read_excel(xl_file, engine="xlrd")
    elif xl_file.find(".xlsb")!=-1:
        # For .xlsb (Excel Binary format):
        df = pd.read_excel(xl_file, engine="pyxlsb")
    else:
        print("invalide extension")
        os._exit(1)

    # Convert to JSON or CSV string
    df = df.rename(columns=lambda x: "" if "Unnamed" in str(x) else x)
    data_str = df.to_csv(csv_file,index=False) 
    # or: data_str = df.to_json(orient="records") jsp


def empty_file(file_name):
    f=open(file_name,"w")
    f.close()


def create_directory(path,folder_name):
    folder_path=path+"/"+folder_name
    if not os.path.exists(folder_path):
        os.makedirs(folder_path, exist_ok=True)
    return folder_path


#---------------------------Other functions----------------------------------


def normalize_to_ddmmyy(date_val: str | int | float | None, delimiter: str = "") -> str:
    """
    Template function to convert any date representation into 'ddmmyy' format.
    
    Args:
        date_val: The date input (ISO string, RFC 2822, timestamp, etc.)
        delimiter: Separator between day, month, year. 
                   Use "" for '250826'
                   Use "/" for '25/08/26'
                   Use "-" for '25-08-26'
    
    Returns:
        Formatted string in ddmmyy. Returns empty string if parsing fails.
    """
    if not date_val:
        return ""

    dt = None

    # 1. Handle Epoch Timestamps (integers or floats)
    if isinstance(date_val, (int, float)):
        try:
            dt = datetime.fromtimestamp(date_val)
        except Exception:
            return ""

    # 2. Handle String Formats
    elif isinstance(date_val, str):
        cleaned_str = date_val.strip()

        # Handle numeric string timestamps (e.g., "1724587200")
        if cleaned_str.isdigit():
            dt = datetime.fromtimestamp(float(cleaned_str))
        else:
            try:
                # Fuzzy parsing handles ISO 8601, RFC 2822, and standard locale dates
                dt = parser.parse(cleaned_str)
            except (ValueError, TypeError):
                return ""

    # 3. Format into ddmmyy (%d = day, %m = month, %y = 2-digit year)
    if dt:
        format_pattern = f"%d{delimiter}%m{delimiter}%y"
        return dt.strftime(format_pattern)

    return ""


def to_bytes(x: str) -> bytes:
    if x.startswith("b'") or x.startswith('b"'):
        return literal_eval(x)
    return x.encode("utf-8")


def clean_email_text(raw_body):
    # 1. Strip out quoted email thread replies (lines starting with '>')
    lines = raw_body.splitlines()
    clean_lines = []

    for line in lines:
        line_str = line.strip()
        # Stop or skip when reply thread history starts
        if line_str.startswith(">") or re.match(
            r"^On .* wrote:$", line_str, re.IGNORECASE
        ):
            break
        clean_lines.append(line_str)

    # 2. Join lines into a single clean text block
    text = " ".join(clean_lines)

    # 3. Clean up unescaped quotes and remaining backslashes
    text = text.replace('\\"', '"').replace("\\", "")

    # 4. Normalize multiple spaces
    return " ".join(text.split())


def day_name(date):
    y=date[:4]
    m=date[5:7]
    d=date[8:10]
    dt=datetime(int(y),int(m),int(d))
    name=dt.strftime("%A")
    DAYS_FR = {
        "Monday": "Lundi",
        "Tuesday": "Mardi",
        "Wednesday": "Mercredi",
        "Thursday": "Jeudi",
        "Friday": "Vendredi",
        "Saturday": "Samedi",
        "Sunday": "Dimanche",
    }
    return DAYS_FR[name]


def saisir(msg):
    x=input("saisir "+msg+": ")
    return x.upper()


#---------------------------Agent Start function---------------------------------


def start_ai_agent(model_name,template_path):
    response=ollama.list()
    ok=False
    for m in response['models']:
        if model_name==m['model']:
            model=OllamaLLM(model=model_name,
                            temperature=0.0,
                            model_kwargs={
                                "seed":42,
                                "top_p":1.0
                            }
            )
            ok=True
    if not ok:
        print("model "+model_name+" is invalide!")
        os._exit()

    if os.path.isfile(template_path):
        template_file=open(template_path,'r')
        template=template_file.read()
        template_file.close()

        prompt = ChatPromptTemplate.from_template(template)
        chain = prompt | model

        return chain
    else:
        print("template path "+template_path+" is not found!")
        os._exit()


#----------------------------Email functions--------------------------


def log_in():
    print("logging in...")
    load_dotenv()

    EMAIL = os.getenv("EMAIL")
    APP_PASSWORD = os.getenv("PASSWORD")
    DOMAINE = os.getenv("DOMAINE")

    # Connect to Gmail
    try:
        mail_obj = imaplib.IMAP4_SSL(DOMAINE)
        mail_obj.login(EMAIL, APP_PASSWORD)
        mail_obj.select("inbox")
    except:
        print("ERROR: connection problem !")
        os._exit(1)

    return mail_obj


def log_out(mail_obj):
    print("logging off...")
    mail_obj.close()
    mail_obj.logout()


def unflag_mail(id,mail_obj):
    mail_obj.store(id,"-FLAGS","\\Flagged")
    print("Email Unflagged!")
    return "Email Unflagged!"


def unsee_mail(id,mail_obj):
    mail_obj.store(id,"-FLAGS","\\SEEN")
    print("Email Unseen!")
    return "Email Unseen!"


def see_mail(id,mail_obj):
    mail_obj.store(id,"+FLAGS","\\SEEN")
    print("Email Seen!")
    return "Email Seen!"


def flag_mail(id,mail_obj):
    mail_obj.store(id,"+FLAGS","\\FLAGGED")
    print("Email Flagged!")
    return "Email Flagged!"


def is_mail_flagged(id ,mail_obj) -> bool:
    """
    Checks if a specific email UID has the '\\Flagged' flag in IMAP.
    
    :param mail_obj: Active imaplib.IMAP4_SSL instance
    :param mail_uid: String representing the IMAP UID of the target email
    :return: True if flagged/starred, False otherwise
    """
    # Fetch flags for the specific email UID
    if isinstance(id, bytes):
        id = id.decode("utf-8")
    else:
        id = str(id)

    status, response = mail_obj.fetch( id, "(FLAGS)")
    
    if status == "OK" and response[0]:
        raw_flags = response[0].decode() if isinstance(response[0], bytes) else str(response[0])
        return "\\Flagged" in raw_flags
    
    return False


def extract_mail_ids(mail_obj):
    result,data = mail_obj.search(None, "UNSEEN")
    ids = data[0].split()
    return ids


def reply_to_mail(i,mail_obj):

    # Fetch original email
    status, data = mail_obj.fetch(i, "(BODY[])")
    raw_msg = email.message_from_bytes(data[0][1])

    sender = raw_msg["From"]
    msg_id = raw_msg["Message-ID"]
    subject = raw_msg["Subject"]

    # --- STEP 2: Use SMTP to send the reply ---
    reply = EmailMessage()
    reply["To"] = sender
    reply["From"] = os.getenv("EMAIL")
    reply["Subject"] = (
    f"Re: {subject}" if not subject.lower().startswith("re:") else subject
    )

    # Set threading headers so Gmail/Outlook group it as a reply
    reply["In-Reply-To"] = msg_id
    reply["References"] = msg_id

    reply.set_content(
    "Hello,\n\nWe received your message and are processing it.\n\nBest regards,"
    )

    # Connect to SMTP server to deliver the mail_obj
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(os.getenv("EMAIL"), os.getenv("PASSWORD"))
        server.send_message(reply)

    print("Reply sent!")


#----------------------Client/Stocker functions---------


def is_valid_client(path,client,stocker):
    l=clients_stockers_pairs(path)
    for cl,st in l:
        if cl==client and st==stocker:
            return True
    return False


def search_mail_order(mid=None,lmid=None,loid=None):
    parametres=read_json_file("parametres.json")

    mails_file_path=os.path.join(parametres["main_directory_path"],"all_mails.jsonl")
    orders_file_path=os.path.join(parametres["main_directory_path"],"orders.jsonl")

    mails=read_jsonl_file(mails_file_path)
    orders=read_jsonl_file(orders_file_path)

    mail={}
    order={}
    i=0
    for mail in mails:
        if mid==None or lmid==None:
            if mail["mail_id"]==mid or mail["local_mail_id"]==lmid:
                mails.pop(i)
                break
        else:
            if mail["mail_id"]==mid and mail["local_mail_id"]==lmid:
                mails.pop(i)
                break
        i+=1

    lmid=mail["local_mail_id"]

    j=0
    for order in orders:
        if order["local_order_id"]==loid or order["local_mail_id"]==lmid:
            orders.pop(j)
            break
        j+=1

    create_jsonl_file(mails,mails_file_path)
    create_jsonl_file(orders,orders_file_path)

    create_json_file(mail,os.path.join(parametres["main_directory_path"],"temp_mail.json"))
    create_json_file(order,os.path.join(parametres["main_directory_path"],"temp_order.json"))

    return mail,order
     

def clients_names_list(path):
    file_path=path+"/Clients.jsonl"
    clients_data=read_jsonl_file(file_path)
    l=[]
    for client in clients_data:
        l.append(client["company"])
    return l


def clients_stockers_pairs(path):
    file_path=path+"/Clients.jsonl"
    clients_data=read_jsonl_file(file_path)
    l=[]
    for client in clients_data:
        for stocker in client["stockeurs"]:
            l.append((client["company"],stocker["stockeur"]))
    return l


def clients_stockers_pairs_text(path):
    l=clients_stockers_pairs(path)
    text=""
    for i in range(len(l)):
        text+="("+l[i][0]+", "+l[i][1]+")"+"\n"
    return text


def add_client(path,client):
    file_path=path+"/Clients.jsonl"

    append_jsonl_file(client,file_path)

    print("Client added successfully!")


def add_stocker(company, path, new_stocker):
    file_path=path+"/Clients.jsonl"
    clients_data=read_jsonl_file(file_path)
    found=False
    i=0
    for client in clients_data:
        if client["company"].upper()==company.upper():
            found=True
            break
        i+=1

    if not found:
        print("Error: company not found!")
        return None

    client["stockeurs"].append(new_stocker)
    clients_data[i]=client
    empty_file(file_path)
    for client in clients_data:
        append_jsonl_file(client,file_path)

    print("Stocker added successfully!")


def stocker_details(company, stocker_name, path):
    file_path=path+"/Clients.jsonl"
    clients_data=read_jsonl_file(file_path)
    company_found=False
    for client in clients_data:
        if client["company"].upper()==company.upper():
            company_found=True
            break
    if not company_found:
        print("Company",company,"is not found!")
        return None,None,None,None
    
    for stockeur in client["stockeurs"]:
        if stockeur["stockeur"].upper()==stocker_name.upper():
            return stockeur["num_stockeur"],stockeur["num_dep"],stockeur["ville"],stockeur["pays"]

    print("Error: Company and stocker couple is invalide ")
    return None,None,None,None


def add_to_order(order_data,stocker_data):
    order_data["num_stockeur"]=stocker_data[0]
    order_data["num_dep"]=stocker_data[1]
    order_data["ville"]=stocker_data[2]
    order_data["pays"]=stocker_data[3]
    return order_data


#------------------------Main functions---------------


def check_temp_files():
    parametres=read_json_file("parametres.json")
    mp=parametres["main_directory_path"]
    mail=read_json_file(os.path.join(mp,"temp_mail.json"))
    order=read_json_file(os.path.join(mp,"temp_order.json"))
    if mail!={}:
        append_jsonl_file(mail,os.path.join(mp,"temp_mail.json"))
        empty_file(os.path.join(mp,"temp_mail.json"))
    if order!={}:
        append_jsonl_file(order,os.path.join(mp,"temp_order.json"))
        empty_file(os.path.join(mp,"temp_order.json"))
    

def update_order(path,data):
    orders_file_path=os.path.join(path,"orders.jsonl")
    mails_file_path=os.path.join(path,"all_mails.jsonl")

    orders_data=read_jsonl_file(orders_file_path)
    mails_data=read_jsonl_file(mails_file_path)

    data["local_order_id"]=int(data["local_order_id"])
    data["local_mail_id"]=int(data["local_mail_id"])
    data["needs_human_interferance"]=False
    is_reg=data.get("is_registered",False)
    if is_reg=="True":
        data["is_registered"]=True
    else:
        data["is_registered"]=False

    is_ord_rel=data.get("is_order_related",False)
    if is_ord_rel=="True":
        data["is_order_related"]=True
    else:
        data["is_order_related"]=False

    for i in range(len(orders_data)):
        if orders_data[i]["local_order_id"]==data["local_order_id"]:
            data["etat"]="Reviewed"
            orders_data[i]=data

    for j in range(len(mails_data)):
        if mails_data[j]["local_mail_id"]==data["local_mail_id"]:
            mails_data[j]["etat"]="Reviewed"
            mail=log_in()
            id=to_bytes(mails_data[j]["mail_id"])
            unsee_mail(id,mail)
            flag_mail(id,mail)
            log_out(mail)

    empty_file(orders_file_path)
    for order in orders_data:
        append_jsonl_file(order,orders_file_path)

    empty_file(mails_file_path)
    for mail in mails_data:
        append_jsonl_file(mail,mails_file_path)

            
def remplir_excel(order_data):
    wb = openpyxl.load_workbook("sage_form_vide_rxl.xlsm",keep_vba=True)

    order_id=order_data["local_order_id"]
    param=read_json_file("parametres.json")
    path=param["orders_directory_path"]+"/"+"Order"+str(order_id)+"/"

    date_data=read_json_file(path+"date"+str(order_id)+".json")

    detail_data=read_jsonl_file(path+"detail"+str(order_id)+".jsonl")

    sheet1=wb["Entête Livraison"]
    cols1=["D","E","F","K","P","R","S"]
    c1="13"
    data1=[
        order_data.get("num_stockeur",None),
        date_data.get("no_vr",None),
        day_name(date_data.get("date","1970-01-01")),
        order_data.get("num_stockeur",None),
        order_data.get("num_dep",None),
        order_data.get("ville",None),
        order_data.get("pays",None)
    ]
    for i in range(len(data1)):
        sheet1[cols1[i]+c1]=data1[i]

    sheet2 = wb["Lignes de Commandes"]
    cols2=["D","E","J"]
    c2=14
    for line in detail_data:
        data2=[
            date_data.get("no_vr",None),
            line.get("Réf SAGEM",None),
            line.get("Qté à commander (en UC)",None)
        ]
        for i in range(len(data2)):
            sheet2[cols2[i]+str(c2)]=data2[i]
        c2+=1

    sheet3 = wb["Int_Livr"]
    cols3_1=["B","C","D","I","N","P","Q"]
    cols3_2=["C","H","M","N","O","P","Q","R","S","T","U"]

    for i in range(len(data1)):
        sheet3[cols3_1[i]+"1"]=data1[i]

    i=2
    for line in detail_data:
        data2=[
            line.get("Réf SAGEM",None),
            line.get("Qté à commander (en UC)",None)
        ]
        for j in range(len(cols3_2)):
            if j<2:
                sheet3[cols3_2[j]+str(i)]=data2[j]
            else:
                sheet3[cols3_2[j]+str(i)]=0
        i+=1

    wb.save(path+"sage_form"+str(order_id)+".xlsm")


def create_data_files(order_id,mail_file,path): 

    parametres=read_json_file("parametres.json")
    chain=start_ai_agent(parametres["model"],parametres["xlsx_reader_template"])

    mail_data=read_json_file(mail_file)
    
    if len(mail_data["Attachments"])==0:
        print("No attachments")
        return None
    csv_file="xl.csv"
    xl_to_csv(mail_data["Attachments"][0],csv_file)
    
    f=open(csv_file,"r",encoding="utf-8",errors="ignore")

    data=""
    line=f.readline()
    while line.find("Designation")==-1 and line!="":
        data+=line
        line=f.readline()

    detail=line
    line=f.readline()
    while line!="":
        detail+=line
        line=f.readline()

    f.close()

    result = chain.invoke({"csv_input":data})

    create_json_file(json_to_dict(result),path+"/"+"date"+str(order_id)+".json")

    f=open(csv_file,"w",encoding="utf-8",errors="ignore")
    f.write(detail)
    f.close()

    f=open(csv_file,"r",encoding="utf-8",errors="ignore")
    reader=csv.DictReader(f)
    for row in reader:
        append_jsonl_file(row,path+"/"+"detail"+str(order_id)+".jsonl")
    f.close()


def mail_extractor(id,mail_obj,mail_file):

    e=dict(mail_id=str,local_mail_id=int,From=str,To=str,Subject=str,Date=str,Body=str,Attachments=list,etat="Pending")
    
    result,msg_data = mail_obj.fetch(id, "(RFC822)")
    raw_email = msg_data[0][1]
    msg = email.message_from_bytes(raw_email)

    e["mail_id"] = id
    parametres=read_json_file("parametres.json")
    e["local_mail_id"]=parametres.get("next_local_mail_id",1)
    parametres["next_local_mail_id"]=e["local_mail_id"]+1
    create_json_file(parametres,"parametres.json")
    e["From"] = msg["From"]
    e["To"] = msg["To"]
    e["Subject"] = msg["Subject"]
    e["Date"] = msg["Date"]
    e["Attachments"]=[]
    e["Body"]=None
    
    # Target MIME types for Excel spreadsheets
    EXCEL_MIME_TYPES = {
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet":"xlsx",  # .xlsx
        "application/vnd.ms-excel.sheet.macroenabled.12":"xlsm",                      # .xlsm
        "application/vnd.ms-excel":"xls"                 # .xls (legacy)
    }

    # Get email body/attachments
    nb_attachments=0
    if msg.is_multipart():
        for part in msg.walk():
            content_type = part.get_content_type()
            if content_type == "text/plain":
                coded_body = part.get_payload(decode=True)
                if coded_body:
                    e["Body"] = coded_body.decode("utf-8", errors="ignore")
            elif content_type in EXCEL_MIME_TYPES:
                filename = part.get_filename()
                if filename:
                    nb_attachments+=1
                    name=parametres["excel_files_directory_path"]+"/"+"fichier"+str(e["local_mail_id"])+"_"+str(nb_attachments) #for review
                    extention = "."+EXCEL_MIME_TYPES[content_type]
                    # Decode payload from base64/quoted-printable to raw binary bytes
                    file_data = part.get_payload(decode=True)
                    if file_data is not None:
                        with open(name+extention, "wb") as f:
                            f.write(file_data)
                        e["Attachments"].append(name+extention)
    else:
        coded_body = msg.get_payload(decode=True)
        if coded_body:
            e["Body"] = coded_body.decode("utf-8", errors="ignore")
    e["Body"]=clean_email_text(e["Body"])
    create_json_file(e,mail_file)
    print("email",id,"has been extracted.")


def analyse_mail(e,chain):
    parametres=read_json_file("parametres.json")
    result = chain.invoke({"pairs":clients_stockers_pairs_text(parametres["main_directory_path"]),"sender":e.get("From",None),"reciever":e.get("To",None),"subject":e.get("Subject",None),"body": e.get("Body",None)})
    data=json_to_dict(result)
    data={"local_mail_id":e["local_mail_id"],**data}
    return data


def process_order(e,mail_obj,data,order_dir_path):
    parametres=read_json_file("parametres.json")
    if is_valid_client(parametres["main_directory_path"],data.get("company","").upper(),data.get("stockeur","").upper()):
        data["is_registered"]=True
        if data["order_type"]=="new_order":
            if len(e["Attachments"])>0:
                create_data_files(data["local_order_id"],"new_mail.json",order_dir_path)
                stocker_data=stocker_details(data.get("company",""),data.get("stockeur",""),parametres["main_directory_path"])
                data=add_to_order(data,stocker_data)
                remplir_excel(data)
                if parametres.get("auto_reply",False)==True:
                    reply_to_mail(e["mail_id"],mail_obj)
                e["etat"]="Done"
                data["etat"]="Ready for WMS"
            else:
                flag_mail(to_bytes(e["mail_id"]),mail_obj)
                e["etat"]="No attachment"
                data["etat"]="No attachment"
                print("No attatchments in mail "+e["mail_id"])
        elif data["order_type"]=="update_order":# not finished
            flag_mail(to_bytes(e["mail_id"]),mail_obj)
            data["etat"]="Not supported order_type"
        elif data["order_type"]=="cancel_order":
            flag_mail(to_bytes(e["mail_id"]),mail_obj)
            data["etat"]="Not supported order_type"
        else:
            flag_mail(to_bytes(e["mail_id"]),mail_obj)
            data["etat"]="Not supported order_type"
    else:
        data["is_registered"]=False
        flag_mail(to_bytes(e["mail_id"]),mail_obj)
        data["etat"]="Lack details"
        e["etat"]="Client Unkwon"
    return e,data


def file_order(e,data,order_directory_path):
    parametres=read_json_file("parametres.json")
    mail_file_path=order_directory_path+"/"+"mail"+str(data["local_mail_id"])+".json"
    order_file_path=order_directory_path+"/"+"order"+str(data["local_order_id"])+".json"

    append_jsonl_file(data,parametres["main_directory_path"]+"/"+"orders.jsonl")

    empty_file(os.path.join(parametres["main_directory_path"],"temp_mail.json"))
    empty_file(os.path.join(parametres["main_directory_path"],"temp_order.json"))

    create_json_file(e,mail_file_path)
    create_json_file(data,order_file_path)


def analyse_order(e, mail_obj, data):

    parametres=read_json_file("parametres.json")

    if data.get("is_order_related",False)==True:

        data={"local_order_id":parametres.get("next_local_order_id",1), **data}
        parametres["next_local_order_id"]=data["local_order_id"]+1
        create_json_file(parametres,"parametres.json")

        order_directory_path=create_directory(parametres["orders_directory_path"],"Order"+str(data["local_order_id"])+"_"+normalize_to_ddmmyy(e["Date"]))+"/"

        if data.get("needs_human_interferance",True)==True:
            data["etat"]="Pending"
            e["etat"]="Must review"
            flag_mail(to_bytes(e["mail_id"]),mail_obj)
        else:
            e,data=process_order(e,mail_obj,data,order_directory_path)

        file_order(e,data,order_directory_path)

    else:
        e["etat"]="Rejected"
        data["etat"]="Not an order"
        append_jsonl_file(data,parametres["main_directory_path"]+"/"+"non_orders.jsonl")
        flag_mail(to_bytes(e["mail_id"]),mail_obj)

    append_jsonl_file(e,parametres["main_directory_path"]+"/"+"all_mails.jsonl")

    print(data.get("etat","No data status"))


#--------------------------Mains-------------------------

#--------------------------Main--------------------------


def main():
    print("RUNNING...")

    parametres=read_json_file("parametres.json")
    
    if parametres.get("model",False)==False:
        print("Error: No model!")
        os._exit(1)
    chain = start_ai_agent(parametres["model"],parametres["mail_reader_template"])

    mail=log_in()

    stop=False
    while not stop:
        print("Listening for emails...")
        print()

        ids=extract_mail_ids(mail)
        print(len(ids),"emails has been found.")
        print()

        if len(ids)>0:
            for id in ids:
                #check_temp_files()
                if is_mail_flagged(id,mail):
                    print("Old mail!")
                    e,data=search_mail_order(mid=str(id))
                    odp=create_directory(parametres["orders_directory_path"],"Order"+str(data["local_order_id"]))
                    e,data=process_order(e,mail,data,odp)
                    file_order(e,data,odp)
                    see_mail(id,mail)
                    if data.get("etat",None)=="Ready for WMS":
                        unflag_mail(id,mail)
                    append_jsonl_file(e,parametres["main_directory_path"]+"/"+"all_mails.jsonl")
                else:
                    print("New mail!")
                    mail_extractor(id,mail,"new_mail.json")
                    e=read_json_file("new_mail.json")
                    data=analyse_mail(e,chain)   
                    analyse_order(e,mail,data)

                print()

                if parametres.get("fast_stop",False)==True:
                    parametres["slow_stop"]=True
                    create_json_file(parametres,"parametres.json")
                    stop=True
                    break

        if parametres.get("slow_stop",True)==True:
            stop=True

        sleep(parametres["delay"])
        parametres=read_json_file("parametres.json")
        print()

    log_out(mail)

    print("END.")


#--------------------------Web Functions-----------------


def affiche_html(text,element):
    pre="<"+element+">"
    post="</"+element+">"
    return pre+text+post


def download():
    if os.path.exists("orders.csv"):
        return send_file(
            "orders.csv",
            mimetype="text/csv",
            as_attachment=True,  # Forces browser to download instead of opening in tab
            download_name="orders.csv",  # Optional: file name the user sees
        )
    else:
        return "<h3>Error: CSV file not found!</h3>", 404


#--------------------------Web Main----------------------


def web_main(): # a reconstruire
    yield affiche_html("RUNNING...","h2")

    parametres=read_json_file("parametres.json")

    if parametres.get("model",False)==False:
        print("Error: No model chosen!")
        os._exit(1)
    chain =start_ai_agent(parametres["model"],parametres["xlsx_reader_template"])

    yield affiche_html("Logging in...","h2")
    mail=log_in()

    while parametres.get("slow_stop",False)==False:
        yield affiche_html("Listening for emails...","h2")
        ids=extract_mail_ids(mail)
        yield affiche_html(str(len(ids))+" emails has been found.","h2")
        if len(ids)>0:
            c=0
            for i in ids:
                c+=1

                mail_extractor(i,mail,"new_mail.json")
                yield affiche_html("email "+str(c)+" has been extracted.","h2")

                e=read_json_file("new_mail.json")

                sender=e.get("From",None)
                reciever=e.get("To",None)
                subject=e.get("Subject",None)
                date_time=e.get("Date",None)
                body=e.get("Body",None)

                result = chain.invoke({"local_order_id":"null","mail_id":e["mail_id"],"sender":sender,"reciever":reciever,"subject":subject,"date_time":date_time,"email_text": body})

                data=json_to_dict(result)
                if data.get("is_order_related",False):
                    parametres=read_json_file("parametres.json")
                    data["local_order_id"]=parametres.get("next_local_order_id",1)
                    parametres["next_local_order_id"]=data["local_order_id"]+1
                    create_json_file(parametres,"parametres.json")
                    append_jsonl_file(data,"orders.jsonl")
                    append_json_to_csv(data,"orders.csv")
                    if data["needs_human_review"]:
                        append_jsonl_file(e,"pending_mails.jsonl")
                        yield affiche_html(flag_mail(i,mail),"h2")
                    else:
                        append_jsonl_file(e,"mails.jsonl")
                        if parametres.get("auto_reply",False):
                            yield affiche_html(reply_to_mail(i),"h2")
                else:
                    append_jsonl_file(data,"non_orders.jsonl")#maindir
                    append_json_to_csv(data,"non_orders.csv")#maindir
                    yield affiche_html(flag_mail(i,mail),"h2")

                yield affiche_html("email analyse finished and data stored in data.jsonl.","h2")
                yield affiche_html("","hr")

        sleep(5)
        parametres=read_json_file("parametres.json")

    log_out(mail)

    yield affiche_html("END.","h2")

    yield affiche_html("<a href='/download'>Download-link</a>","p")

