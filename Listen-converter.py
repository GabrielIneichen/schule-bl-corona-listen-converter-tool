import re
import time
import tkinter as tk
from shutil import copyfile
from tkinter import filedialog

import pandas as pd
from openpyxl import load_workbook

pd.set_option('mode.chained_assignment', None)

# Vorlage und CSV pfad erfragen
root = tk.Tk()
root.withdraw()
print("--------------------------------------------------------------")
print("Ein Tool, geschrieben von Gabriel Ineichen.")
print("Version 1.2 2021-03-23")
print("Lizenz: CC-BY")
print("Kontakt: gabrin@gmx.net\n")
print("Das Programm füllt die Corona-Kontaktliste des Ereignismanagement Basel-Landschaft aus\n"
      "mit Daten einer Klasse, welche aus dem LehrerOffice als CSV exportiert wurden.")
print("Es gibt keine Garantie auf korrekte Funktionalität.")
print("Bitte überprüfen Sie die resultierende Excel-Datei.")
print("\nDas Programm funktioniert am besten, wenn die Voralge des Kantons"
      " und das CSV der Klasse im gleichen Ordner sind.")
print("--------------------------------------------------------------")
print("\n\n")
print("Bitte wähle das Excel File aus, das der Kanton als Vorlage bereitgestellt hat")
vorlage = filedialog.askopenfilename(title="Vorlage des Kantons auswählen",
                                     filetypes=(("Excel files", "*.xlsx"), ("alte Excel files", "*.xls")))
if len(vorlage) < 3:
    quit("Abgebrochen")
print("danke :)")
print()

print("Bitte wähle nun noch das CSV File einer Klasse aus, welches du aus dem LehrerOffice exportiert hast")
file_path = filedialog.askopenfilename(title="Klassenliste auswählen",
                                       filetypes=(("CSV Files", "*.csv"),))
if len(file_path) < 3:
    quit("Abgebrochen")
print("danke :)")
print()

# CSV einlesen und bearbeiten
csv = pd.read_csv(file_path, delimiter=";")
output = csv[
    ["S_Vorname", "S_Name", "S_Geburtsdatum", "S_Geschlecht", "S_Telefon", "S_Mobil", "S_EMail", "S_Strasse", "S_PLZ",
     "S_Ort", "P_ERZ1_Rolle", "P_ERZ1_Name", "P_ERZ1_Vorname"]]
output.loc[:, "S_Geschlecht"] = output.S_Geschlecht.map({"m": "MA", "w": "FE"})
output.loc[:, "Erziehungsberechtigter"] = output.P_ERZ1_Vorname + " " + output.P_ERZ1_Name


# Funktion um die Telefonnummern zu bereinigen
def mapTel(line):
    if str(line.S_Mobil).find("@") != -1 or len(str(line.S_Mobil)) < 8:
        line.S_Mobil = line.S_Telefon
    return line


output = output.apply(mapTel, axis=1)

# Speichern der Daten
cols = ["S_Vorname", "S_Name", "S_Geburtsdatum", "S_Geschlecht", "S_Mobil", "S_Telefon", "S_EMail", "S_Strasse",
        "S_PLZ", "S_Ort"]
cols2 = ["Erziehungsberechtigter"]

path = re.findall(string=vorlage, pattern="(.*\/).*")[0]
file = re.findall(string=file_path, pattern=".*\/(.*\.csv)")[0]
target_excel = path + file.replace(".csv", ".xlsx")
copyfile(vorlage, target_excel)

book = load_workbook(target_excel)
writer = pd.ExcelWriter(target_excel, engine='openpyxl')
writer.book = book
writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
output.to_excel(writer, "Vorlage Ereignismanagement", startrow=2, startcol=0, header=False, index=False, columns=cols)
output.to_excel(writer, "Vorlage Ereignismanagement", startrow=2, startcol=15, header=False, index=False, columns=cols2)

writer.save()
print("Alles erledigt.")
print("Das Excel File mit den SuS ist nun unter " + target_excel + " abgelegt")
print()
print()
print("Ich wünsche dir noch einen wundervollen tag :)")
print("Du kannst das Fenster nun schliessen.")
time.sleep(30)
