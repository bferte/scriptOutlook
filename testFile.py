
#liste des fichiers dans le dossier /dossierRempli

import glob
import pathlib
from pathlib import Path

from threading import Timer


files = []
filesAbsolute = []

for file in glob.glob("dossierRempli/*"):
    files.append(file)

print(files)
print('ici')
for file in files:
    filesAbsolute.append(Path(file).resolve())

for file in filesAbsolute:
    file = str(file) 
    print(file)   
    
print('ici')
print(filesAbsolute)





# Genere l'email via compte outlook local

import win32com.client as win32

outlook = win32.Dispatch('Outlook.Application')
mail = outlook.CreateItem(0)
mail.Display()
mail.To = 'To address'
mail.Subject = 'Message subject'
mail.Body = 'Message body'
mail.HTMLBody = '<h2>HTML Message body</h2>' #this field is optional

# To attach a file to the email (optional):
attachment  = "Path to the attachment"
#mail.Attachments.Add("dossierRempli\pli.txt")
for file in filesAbsolute:
    print(file)
    mail.Attachments.Add(str(file))
#mail.Send()
