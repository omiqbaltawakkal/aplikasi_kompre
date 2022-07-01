from tkinter import *   # use lower case tkinter for python3
from threading import Thread
from time import sleep,time
import socket
import json
import xlsxwriter

header_list = list()
value_list = list()
excel_data = dict()

with open('readme IQBALS.txt', 'r') as d:
    temp = d.read()
    temp = temp.replace("\'", "\"")
    temp = json.loads(temp)
    for key, value in temp.items():
        if excel_data.get(key):
            excel_data.get(key).append(value)
        else:
            temp = list()
            temp.append(value)
            excel_data.update({ key : temp})
    
    # for item in temp:
        # print(item)

with open('readme ASDF.txt', 'r') as f:
    temp = f.read()
    temp = temp.replace("\'", "\"")
    temp = json.loads(temp)
    for key, value in temp.items():
        if excel_data.get(key):
            excel_data.get(key).append(value)
        else:
            temp = list()
            temp.append(value)
            excel_data.update({ key : temp})
print(excel_data)
# workbook = xlsxwriter.Workbook('kkpa.xlsx')
# worksheet = workbook.add_worksheet()

# for x in range(len(header_list)):
#     worksheet.write(0, x, str(header_list[x]))

# for x in range(len(value_list)):
#     for y in range(len(value_list[0])):
#         worksheet.write(1+x, y, str(value_list[x][y]))

# workbook.close()