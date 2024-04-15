import os
import glob
import sys
import json
import csv
import xlwings as xw
import requests
import subprocess
from tqdm import tqdm

import datetime

TODAY = datetime.date.today()
RIGHT_ALIGNMENT = -4152

def get_location(spider):
    result = f"spiders\\{spider}"
    return result

def get_spiders():
    spiders = os.listdir("spiders")
    return spiders

def get_CSV (CSV_URL):
    with requests.Session() as ses:
        download = ses.get(CSV_URL)
        decoded_content = download.content.decode('utf-8', errors='ignore')
        content = list(csv.reader(decoded_content.splitlines() , delimiter = ';'))
    content = content[1:]
    result = [[obiekt.replace(',','.') for obiekt in row]for row in content]
    return result

def fill_TGE_energy():
    DATE = TODAY - datetime.timedelta(days=7)
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

def fill_TGE_gas():
    DATE = TODAY - datetime.timedelta(days=7)
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

def fill_GPI():
    file = open('GPI.json',)
    data = json.load(file)
    file.close()
    excel_energy = xw.Book('ENERGIA.xlsx')
    shutdowns = excel_energy.sheets['Wylaczenia elektrowni']
    shutdowns.range("U7").options(transpose = True).value = shutdowns.range("T7:T9").options(transpose = True).value
    shutdowns.range("D7").value = data[0]["planned"] 
    shutdowns.range("D8").value = data[0]["notplanned"] 
    shutdowns.range("D9").value = data[0]["summaric_leaks"]

def get_dates_to_base(quantity):
    dates = []
    date = datetime.date.today()
    date -= datetime.timedelta(days=7)
    while quantity:
        dates.append(date)
        date += datetime.timedelta(days=1)
        quantity -= 1
    return dates

def fill_BASE(excel_file):
    if excel_file == "GAZ.xlsx":
        data_json = "BASEgas.json"
    elif excel_file == "Energia.xlsx":
        data_json = "BASEenergy.json"
    file = open(data_json)
    data = json.load(file)
    file.close()
    excel = xw.Book(excel_file)
    base = excel.sheets["Kontrakty BASE Y-25 Y-24 Y-23"]
    BASE26DKR = [ i["BASE26-DKR"] for i in data]
    BASE26MWh = [ int(i["BASE26-MWh"]) for i in data]
    BASE25DKR = [ i["BASE25-DKR"] for i in data]
    BASE25MWh = [ int(i["BASE25-MWh"]) for i in data]
    location = int(base.range("H4").value)
    base.range(f"B{location}").options(transpose = True).value = BASE26DKR
    base.range(f"C{location}").options(transpose = True).value = BASE26MWh
    base.range(f"E{location}").options(transpose = True).value = BASE25DKR
    base.range(f"F{location}").options(transpose = True).value = BASE25MWh
    for i in range(len(BASE25DKR)):
        base.range(f"R{location-1}:T{location-1}").copy()
        base.range(f"R{location+i}").paste(paste="formulas")
        base.range(f"D{location-1}").copy()
        base.range(f"D{location+i}").paste(paste="formulas")
        base.range(f"G{location-1}").copy()
        base.range(f"G{location+i}").paste(paste="formulas")
    dates = get_dates_to_base(len(BASE25DKR))
    base.range(f"A{location}").options(transpose = True).value = dates

def get_date_to_url():
    DATE = TODAY - datetime.timedelta(days=1)
    End = DATE.strftime("%Y%m%d")
    DATE -= datetime.timedelta(days=6)
    Begin = DATE.strftime("%Y%m%d")
    DATE += datetime.timedelta(days=7)
    return Begin,End

def fill_cross_border():
    begin , end = get_date_to_url()
    data = get_CSV(f'https://www.pse.pl/getcsv/-/export/csv/PL_WYK_WYM/data_od/{begin}/data_do/{end}')
    excel_energy = xw.Book('ENERGIA.xlsx')
    shutdowns = excel_energy.sheets['Wymiana trans']
    shutdowns.range("A20").value = data

def fill_wind_foto():
    begin, end = get_date_to_url()
    data = get_CSV(f'https://www.pse.pl/getcsv/-/export/csv/PL_GEN_WIATR/data_od/{begin}/data_do/{end}')
    excel_energy = xw.Book('ENERGIA.xlsx')
    generacja = excel_energy.sheets['Generacja wiatr i foto']
    generacja.range("G36").options(transpose = True).value = generacja.range("B36:B37").value
    generacja.range("AI3").value = data

def fill_BASES():
    fill_BASE("GAZ.xlsx")
    fill_BASE("Energia.xlsx")

def remove_JSONS():
    files = glob.glob("*.json")
    for file in files:
        os.remove(file)


if __name__ == "__main__":
    spiders = get_spiders()
    DEBUG = 1
    SCRAPE = 1
    try:
        DEBUG = int(sys.argv[1])
        print(DEBUG)
    except:
        DEBUG = 0
    if SCRAPE and not DEBUG:
        for spider in tqdm(spiders, desc="Scrapping data"):
            spider = get_location(spider)
            subprocess.call(f"python3 {spider}", shell=True)
    if not DEBUG:
        filling = [fill_BASES,fill_wind_foto,fill_cross_border,fill_GPI,fill_TGE_gas,fill_TGE_energy]
    else:
        filling = [fill_GPI]
    for fill in tqdm(filling, desc="Filling excel files"):
        fill()
    remove_JSONS()


