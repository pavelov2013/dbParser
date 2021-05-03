import openpyxl as opx
import os.path
import docx
from datetime import datetime
import time

directory = 'referats/'

files = os.listdir(directory)

wb = opx.load_workbook("data.xlsx")
sheet = wb.active
li = []

for i in range(1,300):
    li.append(str(sheet["E"+str(i+1)].value))

newlist = list()
for i in li:
    for el in files:
        if i in el:
            sheet["G"+str(li.index(i)+2)].value = "+"
            break

wb1 = opx.load_workbook("res.xlsx")
sheet1 = wb1.active
li1 = list()
keys = list()
for i in range(1,545):
    li1.append(str(sheet1["A"+str(i)].value))
    keys.append(str(sheet1["Z"+str(i)].value))
for i in li:
    #l = str("")

    for i1 in li1:
        if i in i1:
            sheet["Z" + str(li.index(i)+2)].value = keys[li1.index(i1)]
            #l = i1
            break
    #sheet["Z" + str(li.index(i)+2)].value = l

wb.save("data.xlsx")
        
