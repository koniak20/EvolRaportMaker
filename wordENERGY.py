import os, sys
from docxtpl import DocxTemplate
import datetime as dt
import xlwings as xw

context = { "prev_week" : "X", "Yprev" : "X", "curr_week" : "X", "Ycurr" : "X", "TGE_ave" : "X", 
       "min_day" : "X", "min_date" : "X" ,"min_price": "X", "max_day" : "X", "max_date" : "X", "max_price": "X",
       "min_hour": "X", "min_hour_price" : "X", "min_hour_date": "X", "min_hour_price" : "X", "max_hour" : "X",
       "max_hour_price" : "X", "max_hour_date" : "X", "max_hour_price" : "X", "per10" : "X", "per25" : "X",
       "per50" : "X", "changeGPI" : "X", "GPI_ave":"X", "GPImin" : "X", "GPImax" : "X", "GPIminP" : "X", 
       "GPImaxP" : "X","GPIminN" : "X", "GPImaxN" : "X", "Cross_ave" : "X", "weather_change" : "X", 
       "weather_max_date" : "X", "weather_max" : "X", "weather_min_date" : "X", "weather_min" : "X",
       "weather_future_change" : "X", "wind_energy_future" : "X", "wind_ave" : "X", "wind_change" : "X",
       "wind_max" : "X", "foto_ave" : "X" ,"foto_change": "X","foto_max" : "X", "wind_future" : "X", "TGEozea_change" :"X",
       "TGEozea_last" : "X", "TGEozea_curr" : "X", "curr_thur" : "X", "TGEozea_change": "X", "TGEozebio_change":"X",
       "TGEozebio_change" : "X", "TGEozebio_curr":"X", "TGEozebio_last":"X","TGeff_change":"X","TGeff_curr":"X",
       "TGeff_last":"X", "last_thru":"X", "TGEozea_MWh":"X", "TGEozebio_MWh":"X","TGeff_MWh":"X",
       "today":"X","BASE23m":10, "BASE23w":11,"BASE24m":10, "BASE24w":11}

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
        week += month + " - " + today.strftime("%d ") + months[today.month - 1]
    else:
        week += " - " + today.strftime("%d ") + month
    return week

def makeDates(TGE):
    context["min_day"] = TGE.range("F10").value
    context["max_day"] = TGE.range("F12").value
    today = dt.date.today()
    context["Ycurr"] = today.year
    today += dt.timedelta(days=-2)
    context["Yprev"] = today.year 
    curr_week = makeWeek(6)
    prev_week = makeWeek(-6)
    context["curr_week"] = curr_week
    context["prev_week"] = prev_week
    context["min_date"] = TGE.range("G10").value.strftime("%d.%m")
    context["max_date"] = TGE.range("G12").value.strftime("%d.%m")
    context["min_hour"] = int(TGE.range("O19").value)
    context["max_hour"] = int(TGE.range("Q19").value)
    context["min_hour_day"] = TGE.range("O16").value
    context["max_hour_day"] = TGE.range("Q16").value
    context["min_hour_date"] = TGE.range("O18").value.strftime("%d.%m")
    context["max_hour_date"] = TGE.range("Q18").value.strftime("%d.%m")
    return today.strftime("%d.%m.%Y")

def putContext(cell): 
    print(cell)
    dict_name, sheet, sheet_place = cell[0] ,cell[1], cell[2]
    times = 1
    rounding = 2
    if sheet_place in ["R2","R3","R6","R5"]:
        times = 100
    context[dict_name] = str(round(sheet.range(sheet_place).value * times,rounding)).replace(".",",")
    if "MWh" in dict_name:
        context[dict_name] = int(float(context[dict_name].replace(",",".")))
def fillTGE(TGE):
    fill = [("TGE_ave",TGE,"F7"),("min_price",TGE,"E10"),("max_price",TGE,"E12"),("per10",TGE,"C16"),("per25",TGE,"D16"),("per50",TGE,"E16"),("min_hour_price",TGE,"O15"),("max_hour_price",TGE,"Q15"),]
    for cell in fill:
        putContext(cell)

def fillBASE(BASE):
    fill = [("BASE24m",BASE,"R5"),("BASE24w",BASE,"R6"),("BASE23m",BASE,"R2"),("BASE23w",BASE,"R3")]
    for cell in fill:
        putContext(cell)

def fillGPI(GPI):
    fill =[("GPI_ave",GPI,"T9"),("GPImin",GPI,"R9"),("GPImax",GPI,"S9"),("GPIminP",GPI,"R7"),("GPImaxP",GPI,"S7"),("GPIminN",GPI,"R8"),("GPImaxN",GPI,"S8")]
    for cell in fill:
        putContext(cell)
    context["changeGPI"] = str(GPI.range("V9").value)


def fillFoto(Foto,Wind):
    fill = [("wind_ave",Foto,"B36"),("wind_max",Foto,"E36"),("foto_ave",Foto,"B37"),("foto_max",Foto,"E37")]
    for cell in fill:
        putContext(cell)
    context["wind_change"] = str(Foto.range("H36").value)
    context["foto_change"] = str(Foto.range("H37").value)
    context["wind_energy_future"] = str(Wind.range("L16").value)
    context["wind_future"] = str(Wind.range("K16").value)


def change(logic_value):
    if logic_value:
        return "wzrosly"
    else :
        return "zmalaly"

def fillTGEozea(Ozea):
    thur = dt.date.today() 
    context['today'] = thur.strftime("%d.%m.%Y")
    thur -= dt.timedelta(days=4)
    context['curr_thur'] = thur.strftime("%d.%m")
    thur -= dt.timedelta(days=7)
    context['last_thur'] = thur.strftime("%d.%m")
    context['TGEozea_change'] = change(Ozea.range("H4").value)
    context['TGEozebio_change'] = change(Ozea.range("J4").value)
    context['TGeff_change'] = change(Ozea.range("L4").value)
    fill = [("TGEozea_last",Ozea,"H3"),("TGEozea_curr",Ozea,"H2"),("TGEozebio_curr",Ozea,"J2"),("TGEozebio_last",Ozea,"J3"),("TGeff_curr",Ozea,"L2"),("TGeff_last",Ozea,"L3"),("TGEozea_MWh",Ozea,"I2"),("TGEozebio_MWh",Ozea,"K2"),("TGeff_MWh",Ozea,"M2")]
    for cell in fill:
       putContext(cell)


if "__main__" == __name__:
    os.chdir(sys.path[0])

    energy = xw.Book('Energia.xlsx')
    TGE = energy.sheets['TGE RDN']
    BASE = energy.sheets['Kontrakty BASE Y-25 Y-24 Y-23']
    GPI = energy.sheets['Wylaczenia elektrowni']
    Foto = energy.sheets['Generacja wiatr i foto']
    Cross = energy.sheets['Wymiana trans']
    Ozea = energy.sheets['OZE zielone']
    Wind = energy.sheets['Prognoza wiatr']
    context["Cross_ave"] = str(round(Cross.range("B17").value,2)).replace(".",",")
    fillTGE(TGE)
    makeDates(TGE)
    fillBASE(BASE)
    fillGPI(GPI)
    fillFoto(Foto,Wind)
    fillTGEozea(Ozea)
    tod = dt.date.today()
    doc = DocxTemplate('ENERGIAtem.docx')
    doc.render(context)
    doc.save(f'ENERGIA {tod.strftime("%d.%m.%Y")}.docx')
