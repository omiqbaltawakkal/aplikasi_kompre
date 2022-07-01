from tkinter import *   # use lower case tkinter for python3
from threading import Thread
from time import sleep,time
import socket


with open('readme '+worker_name_entry.get()+'.txt', 'r') as f:
    temp = f.read()
    
    workbook = xlsxwriter.Workbook('kkpa.xlsx')
    worksheet = workbook.add_worksheet()

    for x in range(len(header_list)):
        worksheet.write(0, x, str(header_list[x]))

    for x in range(len(collection_data_value_only)):
        for y in range(len(collection_data_value_only[0])):
            worksheet.write(1+x, y, str(collection_data_value_only[x][y]))

    workbook.close()