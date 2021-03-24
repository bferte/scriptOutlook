
#liste des fichiers dans le dossier /dossierRempli

import glob
import pathlib
from pathlib import Path

import win32com.client as win32

import pythoncom



# watchdog
import sys
import time
import logging
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

def on_created(event):
    print("created")
    files = []
    filesAbsolute = []

    for file in glob.glob("dossierRempli/*"):
        files.append(file)


    for file in files:
        filesAbsolute.append(Path(file).resolve())

    for file in filesAbsolute:
        file = str(file) 
        print(file)   
        

    # Genere l'email via compte outlook local
    pythoncom.CoInitialize()




    outlook = win32.Dispatch('Outlook.Application')
    mail = outlook.CreateItem(0)
    mail.Display()
    #mail.To = 'To address'
    #mail.Subject = 'Message subject'
    #mail.Body = 'Message body'
    #mail.HTMLBody = '<h2>HTML Message body</h2>' #this field is optional

    # To attach a file to the email (optional):
    #attachment  = "Path to the attachment"
    #mail.Attachments.Add("dossierRempli\pli.txt")
    for file in filesAbsolute:
        mail.Attachments.Add(str(file))
        

def on_moved(event):
    print("moved")




#from threading import Timer
if __name__ == "__main__":
    
    event_handler = FileSystemEventHandler()

    #calling functions
    event_handler.on_created = on_created
    event_handler.on_moved = on_moved

    path="C:/Users/inove/OneDrive/Bureau/DEV Briac/scriptOutlook/dossierRempli"
    observer = Observer()
    observer.schedule(event_handler, path, recursive=True)

    observer.start()
    try:
        print("Surveillance")
        while True:
            time.sleep(1)
    finally:
        observer.stop()
        print('done')
    observer.join()

