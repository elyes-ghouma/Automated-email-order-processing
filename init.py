import os

f=open("data_files.txt","r")
ch=f.readline()[:-1]
while ch!="":
    if os.path.exists(ch):
        os.remove(ch)
        print("file deleted.")
    else:
        print("file not found.")
    ch=f.readline()[:-1]

f.close()
