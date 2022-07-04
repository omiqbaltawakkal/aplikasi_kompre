from tkinter import *   # use lower case tkinter for python3
from threading import Thread
from time import sleep,time
import socket
import xlsxwriter
import json
import subprocess

def callPowershellFunc(cmd):
    result = subprocess.Popen(
        ["powershell", "-Command", cmd], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    out, err = result.communicate()
    return out

f = open('readme_self.excel_collection_value.txt', 'r')
data = f.read()
data = json.loads(data.replace("\'", "\""))
# print(type(data))
list_temp = list()
for each in data :
    temp = dict()
    for key, value in each.items():
        if type(value) == list:
            strings = ''
            for item in value:
                if type(item)== list:
                    strings +=','.join(item)
                    strings += "\r\n"
                    value = strings
                elif type(item) == dict:
                    strings +=','.join('='.join((key,val)) for (key,val) in item.items())
                    strings += "\r\n"
                    value = strings
        else:
            pass
        # print(key, value)
        temp.update({key: value})
    list_temp.append(temp)
    
print (list_temp)
    

# print(data)

# result = callPowershellFunc(
#         "Get-WmiObject win32_networkadapterconfiguration | Select-Object -Property @{Name = 'IPAddress' ; Expression = {($PSItem.IPAddress[0])}}, MacAddress | Where IPAddress -NE $null | ft -HideTableHeaders")
#     #Get-WmiObject Win32_NetworkAdapterConfiguration -Filter "DHCPEnabled=$True" | Where-Object {$_.IPEnabled -AND $_.IPAddress -gt 0} |Select-object IPAddress, MACAddress
# result = result.decode("utf-8").strip().split("\r\n")
# result = [each.split(" ") for each in result]

# print (result)

# workbook = xlsxwriter.Workbook('kkpa.xlsx')
# worksheet = workbook.add_worksheet()
# for x in range(len(header)):
    # print(header[x])
    # worksheet.write(0, x, str(header[x]))

# for x in range(len(values)):
#     for y in range(len(values[0])):
#         worksheet.write(1+x, y, str(values[x][y]))

# workbook.close()