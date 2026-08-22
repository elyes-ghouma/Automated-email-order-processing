from functions_real import *

mail=log_in()

param=read_json_file("parametres.json")

chain=start_ai_agent(param["model"],param["mail_reader_template"])
ch=str(b'7343')
print(ch.encode("utf-8"))
flag_mail(bytes(ch),mail)