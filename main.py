import os
import glob
import sys
import json
import csv
import xlwings as xw
import requests
import platform
import subprocess
from tqdm import tqdm

import datetime

DATEg = datetime.date.today()
RIGHT_ALIGNMENT = -4152

def getLocation(isUnix, spider):
    if isUnix:
        result = f"/home/koniak20/Desktop/Raport/EvolutionRaport/spiders/{spider}"
    else:
        result = f"spiders\\{spider}"
    return result

def getSpiders(isUnix):
    if isUnix:
        spiders = subprocess.check_output("ls /home/koniak20/Desktop/Raport/EvolutionRaport/spiders", shell=True).decode("utf-8")
    else:
        spiders = subprocess.run(["powershell", "-Command","ls spiders | select name | ft -hide"],shell=True, stdout=subprocess.PIPE).stdout.decode("utf-8").replace("\r","").replace(" ","")
    result = spiders.split("\n")
    result = list(filter(None, result))
    return result

def checkSystem():
    return platform.system() != "Windows"

def getCSV (CSV_URL):
    with requests.Session() as ses:
        download = ses.get(CSV_URL)
        decoded_content = download.content.decode('utf-8', errors='ignore')
        content = list(csv.reader(decoded_content.splitlines() , delimiter = ';'))
    content = content[1:]
    result = [[obiekt.replace(',','.') for obiekt in row]for row in content]
    return result

def fillTGEenergy ():
    DATE = DATEg - datetime.timedelta(days=7)
    file = open('TGEenergy.json',)
    data = json.load(file)
    file.close()
    price = []
    hours = []
    for lista in data:
        price.append((lista["price"])) #NEW JSON NAMES
        hours.append([(lista["min price"]), lista["hour for min"],lista["max price"],lista["hour for min"]])
    excel_energy = xw.Book('ENERGIA.xlsx')
    TGE = excel_energy.sheets['TGE RDN']
    values = list(map(lambda x : round(x,2),filter(lambda y : (y is not None),(TGE.range("B16:B400").value))))
    TGE.range("C5").value = DATE
    TGE.range("B17").options(transpose = True).value = values
    TGE.range("B6").options(transpose = True).value = price
    TGE.range("O6").value = hours

def fillTGEgas():
    DATE = DATEg - datetime.timedelta(days=7)
    file = open('TGEgas.json',)
    data = json.load(file)
    file.close()
    price = []
    for row in data:
        price.append((row["price"])) 
    excel_energy = xw.Book('GAZ.xlsx')
    TGE = excel_energy.sheets['TGE']
    values = list(map(lambda x : round(x,2),filter(lambda y : (y is not None),(TGE.range("B16:B400").value))))
    TGE.range("B5").value = DATE
    TGE.range("B17").options(transpose = True).value = values
    TGE.range("C7").options(transpose = True).value = price

def fillGPI():
    file = open('GPI.json',)
    data = json.load(file)
    file.close()
    excel_energy = xw.Book('ENERGIA.xlsx')
    Wylaczenia = excel_energy.sheets['Wylaczenia elektrowni']
    Wylaczenia.range("U7").options(transpose = True).value = Wylaczenia.range("T7:T9").options(transpose = True).value
    Wylaczenia.range("D7").value = data[0]["planned"] 
    Wylaczenia.range("D8").value = data[0]["notplanned"] 
    Wylaczenia.range("D9").value = data[0]["summaric_leaks"]

def getDatesToBASE(quantity):
    dates = []
    date = datetime.date.today()
    date -= datetime.timedelta(days=7)
    while quantity:
        dates.append(date)
        date += datetime.timedelta(days=1)
        quantity -= 1
    return dates

def fillBASE(excel_file):
    if excel_file == "GAZ.xlsx":
        datajson = "BASEgas.json"
    elif excel_file == "Energia.xlsx":
        datajson = "BASEenergy.json"
    file = open(datajson)
    data = json.load(file)
    file.close()
    excel = xw.Book(excel_file)
    base = excel.sheets["Kontrakty BASE Y-25 Y-24 Y-23"]
    BASE25DKR = [ i["BASE25-DKR"] for i in data]
    BASE25MWh = [ int(i["BASE25-MWh"]) for i in data]
    BASE24DKR = [ i["BASE24-DKR"] for i in data]
    BASE24MWh = [ int(i["BASE24-MWh"]) for i in data]
    location = int(base.range("H4").value)
    base.range(f"B{location}").options(transpose = True).value = BASE25DKR
    base.range(f"C{location}").options(transpose = True).value = BASE25MWh
    base.range(f"E{location}").options(transpose = True).value = BASE24DKR
    base.range(f"F{location}").options(transpose = True).value = BASE24MWh
    for i in range(len(BASE25DKR)):
        base.range(f"O{location-1}:Q{location-1}").copy()
        base.range(f"O{location+i}").paste(paste="formulas")
        base.range(f"D{location-1}").copy()
        base.range(f"D{location+i}").paste(paste="formulas")
        base.range(f"G{location-1}").copy()
        base.range(f"G{location+i}").paste(paste="formulas")
    dates = getDatesToBASE(len(BASE24DKR))
    base.range(f"A{location}").options(transpose = True).value = dates

def DateURL():
    DATE = DATEg
    DATE -= datetime.timedelta(days=1)
    End = DATE.strftime("%Y%m%d")
    DATE -= datetime.timedelta(days=6)
    Begin = DATE.strftime("%Y%m%d")
    DATE += datetime.timedelta(days=7)
    return Begin,End

def fillCrossBorder():
    begin , end = DateURL()
    data = getCSV(f'https://www.pse.pl/getcsv/-/export/csv/PL_WYK_WYM/data_od/{begin}/data_do/{end}')
    excel_energy = xw.Book('ENERGIA.xlsx')
    Wylaczenia = excel_energy.sheets['Wymiana trans']
    Wylaczenia.range("A20").value = data

def fillWindFoto():
    begin, end = DateURL()
    data = getCSV(f'https://www.pse.pl/getcsv/-/export/csv/PL_GEN_WIATR/data_od/{begin}/data_do/{end}')
    excel_energy = xw.Book('ENERGIA.xlsx')
    generacja = excel_energy.sheets['Generacja wiatr i foto']
    generacja.range("G36").options(transpose = True).value = generacja.range("B36:B37").value
    generacja.range("AI3").value = data

def fillB():
    fillBASE("GAZ.xlsx")
    fillBASE("Energia.xlsx")

def removeJSONS():
    files = glob.glob("*.json")
    for file in files:
        os.remove(file)


if __name__ == "__main__":
    isUnix = checkSystem() #Created only for unix and windows
    spiders = getSpiders(isUnix)
    DEBUG = 1
    SCRAPE = 1
    try:

        DEBUG = int(sys.argv[1])
        print(DEBUG)
    except:
        DEBUG = 0
    if SCRAPE and not DEBUG:
        for spider in tqdm(spiders, desc="Scrapping your data"):
            spider = getLocation(isUnix,spider)
            subprocess.call(f"python3 {spider}", shell=True)
    if not DEBUG:
        filling = [fillB,fillWindFoto,fillCrossBorder,fillGPI,fillTGEgas,fillTGEenergy]
    else:
        filling = [fillGPI]
    for fill in tqdm(filling, desc="Filling your excel files"):
        fill()
    removeJSONS()


