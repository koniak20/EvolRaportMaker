# -*- coding: utf-8 -*-
import os, sys
from docxtpl import DocxTemplate
import datetime as dt
import xlwings as xw

context = {'previous_week': 'DO',
                'curr_week': 'DO', 'year':2, 'TGE_ave':3,
                'min_day' : 4, 'max_day': 5, 'min_price':6,
                'max_price': 7, 'per10': 8, 'per25':9, 'per50':10,
                'BASE23m':10, 'BASE23w':11,'BASE24m':10, 'BASE24w':11}

def putContext(cell): 
    dict_name, sheet, sheet_place = cell[0] ,cell[1], cell[2]
    times = 1
    if sheet_place in ["R2","R3","R6","R7"]:
        times = 100
    context[dict_name] = str(round(sheet.range(sheet_place).value * times,2)).replace(".",",")

def makeWeek(dayshift):
    months = ["stycznia","lutego","marca","kwietnia","maja","czerwca","lipca","sierpnia","wrzesnia","pazdziernika","listopada","grudnia"]
    today = dt.date.today()
    if dayshift < 0:
        today -= dt.timedelta(days=7)
        dayshift *= -1
    week = today.strftime("%d ")
    month = months[today.month-1]
    today += dt.timedelta(days=dayshift)
    if month != months[today.month-1]:
        week += month + "- " + today.strftime("%d ") + months[today.month - 1]
    else:
        week += "- " + today.strftime("%d ") + month
    return week

def makeDates(TGE):
    context["min_day"] = TGE.range("G9").value
    context["max_day"] = TGE.range("G11").value
    today = dt.date.today()
    context["year"] = today.year 
    curr_week = makeWeek(6)
    prev_week = makeWeek(-6)
    context["curr_week"] = curr_week
    context["previous_week"] = prev_week
    return today.strftime("%d.%m.%Y")

if __name__ == "__main__":
    os.chdir(sys.path[0])
    doc = DocxTemplate('GAZtem.docx')
    gas = xw.Book('GAZ.xlsx')
    TGE = gas.sheets['TGE']
    BASE = gas.sheets['Kontrakty BASE Y-25 Y-24 Y-23']
    fill = [("TGE_ave",TGE,"F8"),("min_price",TGE,"F9"),("max_price",TGE,"F11"),("per10",TGE,"C16"),("per25",TGE,"D16"),("per50",TGE,"E16"),("BASE24m",BASE,"R6"),("BASE24w",BASE,"R7"),("BASE23m",BASE,"R2"),("BASE23w",BASE,"R3")]
    for cell in fill:
        putContext(cell)
    data = makeDates(TGE)
    doc.render(context)
    doc.save(f'GAZ {data}.docx')
