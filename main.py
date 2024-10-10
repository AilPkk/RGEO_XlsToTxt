# This script reads text from catalogue and creates pa.txt files

from openpyxl import load_workbook
from pathlib import WindowsPath
from tkinter import Tk
from tkinter.filedialog import askdirectory
from datetime import datetime
import os

# Ask for folder and generate full file and folder paths
work_folder_path = askdirectory(title='Vali kaust')
xls_path = work_folder_path+"\\Kataloog.xlsx\""
out_path = work_folder_path+"\\pa\\\""
xls_path = WindowsPath(xls_path.replace('"', ''))
os.makedirs(work_folder_path+"/pa", exist_ok=True)

#load workbook sheets as tables
workbook = load_workbook(xls_path, read_only=True, data_only=True)
TULP = []
KIHID = []

sheet = workbook["TULP"]
for row in sheet.values:
    row = ['' if x is None else x for x in row]
    try:
        row[6] = row[6].strftime("%d.%m.%Y")
        row[8] = row[8].strftime("%d.%m.%Y")
    except:
        pass # Bad stuff
    TULP.append(list(row[2:]))

sheet = workbook["KIHID"] #There is a more elaborate way of doing it, but I'm in hurry
for row in sheet.values:
    row = ['' if x is None else x for x in row]
    del row[4]
    KIHID.append(list(row))

workbook.close()

#parse the tables and output pa.txt
# add header
for row in TULP[1:]:
    if row[1] != "":
        print_list = ['**TULP\n']
        pa_name = row[0]
        print_list.append('Uuringu punkt:%s\n' % str(row[0]) +
                          'Maapinna kõrgus:%s\n' % str(row[1]) +
                          'Uuringu sügavus:%s\n' % str(row[2]) +
                          'Seade:%s\n' % str(row[3]) +
                          'Uuringu kuupäev:%s\n' % str(row[4]) +
                          'Pinnasevee sügavus:%s\n' % str(row[5]) +
                          'Pinasevee kuupäev:%s\n' % str(row[6]) +
                          'X koordinaat:X=%s\n' % str(row[7]) +
                          'Y koordinaat:Y=%s\n' % str(row[8]) +
                          'Pikett:%s\n' % str(row[9]) +
                          'Asukoht tee telje suhtes:%s\n' % str(row[10]) +
                          '**KIHID\n'
                          )

# add layers
        for kiht in KIHID:
            if kiht[0] == pa_name:
                print_list.append('Kiht:%s\n' % str(kiht[1]) +
                                  'Algus:%s\n' % str(kiht[2]) +
                                  'Kirjeldus:%s\n' % str(kiht[4]) +
                                  'Geoindeks:%s\n' % str(kiht[3])
                                  )
        print_list.append('**LABOR\n**Lopp')

        file_path = out_path+pa_name+".txt"
        file_path = WindowsPath(file_path.replace('"', ''))
        with open(file_path, "w", encoding='ANSI') as f:
            f.write("".join(print_list))
