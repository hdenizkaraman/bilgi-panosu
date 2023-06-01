import os
from bs4 import BeautifulSoup
import requests
import re
import openpyxl
import locale
import datetime

denq = """
██████╗░███████╗███╗░░██╗░██████╗░
██╔══██╗██╔════╝████╗░██║██╔═══██╗
██║░░██║█████╗░░██╔██╗██║██║██╗██║
██║░░██║██╔══╝░░██║╚████║╚██████╔╝
██████╔╝███████╗██║░╚███║░╚═██╔═╝░
╚═════╝░╚══════╝╚═╝░░╚══╝░░░╚═╝░░░
"""

def teacherlist():
    wb = openpyxl.load_workbook('nobetciler.xlsx')
    sheet = wb['Sheet1']
    excel_days = ["C1","E1","G1","I1","K1","M1","O1"]
    locale.setlocale(locale.LC_ALL, '')
    nowdate = datetime.datetime.now()
    todayname = datetime.datetime.strftime(nowdate, '%A')
    
    for day in excel_days:
        getdayname = sheet[day].value
        if getdayname=="cumartesi" or getdayname=="pazar":
            print("Bugün Haftasonu Dostum. Git ve Kahveni İç...")
            columnofday = "None"
        if todayname.lower() == getdayname.lower():
            columnofday = day[0]
            break
    
        
    dayguardian = [str(columnofday)+str(i) for i in range(3,15,2)]
    floorlist = ["A"+str(i) for i in range(3,15,2)]

    codelist = []
    for guard, floor in zip(dayguardian, floorlist):
        thecode = f"<tr> <th scope='row'>{sheet[floor].value}</th> <td>{sheet[guard].value}</td> </tr>"
        codelist.append(thecode)

    thecode_str = "\n".join(map(str, codelist))
    code_space = soup.find("div", {"id":"PYTHON_NOBETCI_LISTESI"})
    code_space.string = thecode_str

    


def sliderImages():
    imagelist = os.listdir(imagefolder)
    codelist = []
    for imagename in imagelist:
        imagecode = f"<div class='carousel-item'><img src='{imagefolder+imagename}' class='scale' data-scale='best-fit' data-align='center'></div>"
        codelist.append(imagecode)
    slidercode = "\n".join(map(str, codelist))
    image_space = soup.find("div", {"id":"PYTHON_SLIDER"})
    image_space.string = slidercode
    


projectfolder = os.path.dirname(os.path.realpath(__file__))
template = projectfolder+"\\arayuz.html"
imagefolder = "static/img/"

print(denq)
site = open(template, encoding="utf8")
soup = BeautifulSoup(site, "html.parser")
teacherlist()
sliderImages()



# Günlük Dosyayı Kaydet
with open("index.html", "w", encoding = 'utf-8') as file:
    file.write(str(soup.prettify(formatter=None)))

os.system("index.html")
