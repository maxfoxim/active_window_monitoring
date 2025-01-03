from typing import Optional
from ctypes import wintypes, windll, create_unicode_buffer
import time
import json
import matplotlib.pyplot as plt
from datetime import datetime

"""
TODO:
- Minuten statt Sekunden
- sonstige Kategorie
"""
SCHLAFZEIT=5 # Abstand Messung
ANZAHL_ITERATIONEN= 10**10
ALLE_SEKUNDEN_SPEICHERN_JSON=5*60/SCHLAFZEIT
ALLE_SEKUNDEN_SPEICHERN_BARCHART=60*5/SCHLAFZEIT
MINUTEN=60 # Umrechnung Sekunden in Minuten
MINDESTDAUER = 60 # Mindestdauer Fenster offen für Zeigen in Chart

Programme = ["Visual Studio Code", "VMware Remote Console","Excel",
             "DBeaver", "Mozilla Firefox", "Rainbow","Windows PowerShell",
             "Outlook","MaInSpec Validator","Zeiss Configuration Creator", 
             "OneNote","Microsoft​ Edge"]

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
letztes_datum=""
# mit 0 initalisiern
for Programm in Programme:
    anwendung[Programm]=0

for i in range(ANZAHL_ITERATIONEN):
    print("------------------------------")
    aktuelles_fenster=getForegroundWindowTitle()
    
    # Bereinigung:
    try:
        aktuelles_fenster=aktuelles_fenster.replace("● ","")
    except:
        pass
    
    now = datetime.now() # current date and time
    aktuelles_datum = now.strftime("%m_%d")

    # mit 0 initalisiern (auch bei neuem Tag)
    if aktuelles_datum != letztes_datum:
        for Programm in Programme:
            anwendung[Programm]=0  
        fenster={}
        
    letztes_datum = aktuelles_datum
    
    try:

        # Genaues Fenster
        if aktuelles_fenster in fenster:
            fenster[aktuelles_fenster]=fenster[aktuelles_fenster]+SCHLAFZEIT
        else:
            fenster[aktuelles_fenster]=SCHLAFZEIT

        # Anwendung allgemein
        for Programm in Programme:
            if Programm in aktuelles_fenster:
                anwendung[Programm]=anwendung[Programm]+SCHLAFZEIT
                print("Aktuelle Anwendung:",Programm)
    except:
         print("Konnte nicht gespeichert werden.")
            
    print(aktuelles_fenster)
    time.sleep(SCHLAFZEIT)
    
    if i%ALLE_SEKUNDEN_SPEICHERN_JSON==0:
        with open('Auswertung_Offene_Fenster.json', 'w', encoding ='utf8') as json_file:
            gesamt={"Fenster":fenster,"Anwendung":anwendung}
            json.dump(gesamt, json_file, ensure_ascii = False, indent=4)
    

    # Plotte Fenster 
    if i%ALLE_SEKUNDEN_SPEICHERN_BARCHART==0:
        x_arr,y_arr=[],[]
        for key in fenster:
            if (type(key)==str) and (MINDESTDAUER < fenster[key]):
                try:
                    x_arr.append(key)
                    y_arr.append(round(fenster[key]/MINUTEN,1))
                except:
                    print("Nicht speicherbar")
                
        x_arr = [x for _,x in sorted(zip(y_arr,x_arr))]
        y_arr = sorted(y_arr)
            
        print(x_arr,y_arr)
        plt.clf()
        plt.figure(figsize=(18,10))        
        plt.barh(x_arr,y_arr)
        for x,y in zip(x_arr,y_arr):
            plt.text(y,x,y,fontsize=16,horizontalalignment='left',verticalalignment='center')
        plt.title("Dauer Fenster")
        plt.subplots_adjust(bottom=0.05,top=0.95,left=0.3,right=0.99)
        plt.savefig("OffeneFenster_Details.png")
        
        plt.savefig("Offene_Fenster/Programm_"+aktuelles_datum+"_Details.png")
        plt.clf()
        plt.close()



        # Plotte Anwendungen
        x_arr,y_arr=[],[]
        for key in anwendung:
            if (anwendung[key]>0) and  (type(key)==str) and (MINDESTDAUER < anwendung[key]):
                x_arr.append(key)
                y_arr.append(round(anwendung[key]/MINUTEN,1))
        
        x_arr = [x for _,x in sorted(zip(y_arr,x_arr))]
        y_arr = sorted(y_arr)
        
        x_arr=x_arr[::-1]
        y_arr=y_arr[::-1]

        plt.clf()
        plt.figure(figsize=(12,8))
        plt.bar(x_arr,y_arr)
        plt.xticks(rotation=45)
        #plt.grid()
        for x,y in zip(x_arr,y_arr):
            plt.text(x,y,y,fontsize=16,horizontalalignment='center',verticalalignment='bottom')
        plt.title("Dauer Programm")
        plt.subplots_adjust(bottom=0.2,left=0.05,right=0.95)
        plt.savefig("OffeneFenster.png")
        
        plt.savefig("Offene_Fenster/Programm_"+aktuelles_datum+".png")
        plt.clf()
        plt.close()
