from typing import Optional
from ctypes import wintypes, windll, create_unicode_buffer
import time
import json
import matplotlib.pyplot as plt
from datetime import datetime

"""
TODO:
- Minuten statt Sekunden
- Einzelanwendungen auch plotten

"""

ANZAHL_ITERATIONEN=6 #* 10^10
ALLE_SEKUNDEN_SPEICHERN_JSON=5
SCHLAFZEIT=2 #
ALLE_SEKUNDEN_SPEICHERN_BARCHART=60*5

Programme = ["Visual Studio Code", "VMware Remote Console",
             "DBeaver", "Mozilla Firefox", "Rainbow","Windows PowerShell","Outlook"]

def getForegroundWindowTitle() -> Optional[str]:
    hWnd = windll.user32.GetForegroundWindow()
    length = windll.user32.GetWindowTextLengthW(hWnd)
    buf = create_unicode_buffer(length + 1)
    windll.user32.GetWindowTextW(hWnd, buf, length + 1)
    
    # 1-liner alternative: return buf.value if buf.value else None
    if buf.value:
        return buf.value
    else:
        return None

fenster={}
anwendung={}
for Programm in Programme:
    anwendung[Programm]=0

for i in range(ANZAHL_ITERATIONEN):
    print("------------------------------")
    aktuelles_fenster=getForegroundWindowTitle()
    
    try:
        # Genaues Fenster
        if aktuelles_fenster in fenster:
            fenster[aktuelles_fenster]=fenster[aktuelles_fenster]+1
        else:
            fenster[aktuelles_fenster]=1

        # Anwendung allgemein
        for Programm in Programme:
            if Programm in aktuelles_fenster:
                anwendung[Programm]=anwendung[Programm]+1
                print("Aktuelle Anwendung:",Programm)
    except:
         print("Konnte nicht gespeichert werden.")
            
    print(aktuelles_fenster)
    time.sleep(SCHLAFZEIT)
    
    if i%ALLE_SEKUNDEN_SPEICHERN_JSON==0:
        with open('Auswertung_Offene_Fenster.json', 'w', encoding ='utf8') as json_file:
            gesamt={"Fenster":fenster,"Anwendung":anwendung}
            json.dump(gesamt, json_file, ensure_ascii = False, indent=4)
    
    if i%ALLE_SEKUNDEN_SPEICHERN_BARCHART==0:
        x_arr,y_arr=[],[]
        for key in anwendung:
            x_arr.append(key)
            y_arr.append(anwendung[key])
        plt.clf()
        plt.figure(figsize=(12,8))
        plt.bar(x_arr,y_arr)
        plt.xticks(rotation=45)
        #plt.grid()
        plt.title("Dauer Programm")
        plt.subplots_adjust(bottom=0.2,left=0.05,right=0.95)
        plt.savefig("Test.png")
        
        now = datetime.now() # current date and time
        folder = now.strftime("%m_%d")
        plt.savefig("Offene_Fenster/Programm_"+folder+".png")

print(anwendung)
print(fenster)