from tkinter import *   # use lower case tkinter for python3
from threading import Thread
from time import sleep, time
import socket
import xlsxwriter
import json
import subprocess


def callPowershellFunc(cmd):
    result = subprocess.Popen(
        ["powershell", "-Command", cmd], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    out, err = result.communicate()
    return out


result = callPowershellFunc(
    "Get-WmiObject win32_networkadapterconfiguration | Select-Object -Property @{Name = 'IPAddress' ; Expression = {($PSItem.IPAddress[0])}}, MacAddress | Where IPAddress -NE $null | ft -HideTableHeaders")
   # Get-WmiObject Win32_NetworkAdapterConfiguration -Filter "DHCPEnabled=$True" | Where-Object {$_.IPEnabled -AND $_.IPAddress -gt 0} |Select-object IPAddress, MACAddress
if result:
    result = result.decode("utf-8").strip().split("\r\n")
    print(result[0].split()[0] + " / "+ result[0].split()[1])
else:
    report_collection.update({collection_key: "IP Not Found"})


# storage_capacity = callPowershellFunc(
#     "Get-WmiObject win32_logicaldisk | Format-Table @{label=\"Nama\";e={$_.DeviceId} }, @{n=\"Size\";e={\'{0}GB\' -f [math]::Round($_.Size/1GB,2)}},@{n=\"FreeSpace\";e={\'{0}GB\' -f [math]::Round($_.FreeSpace/1GB,2)}}")
# storage_capacity = list(
#     filter(None, storage_capacity.decode("utf-8").strip().split("\r\n")))
# del storage_capacity[1]  # delete ---- ----- ---- ----
# # [DeviceId, DriveType, Size, FreeSpace ]
# item_dict_keys = storage_capacity[0].split()
# del storage_capacity[0]  # delete header
# print(storage_capacity)
# storage_list = list()
# increment_size = 5
# for x in range(0, len(storage_capacity)):
#     item_dict = dict()
#     storage_item = list(filter(None, storage_capacity[x].split()))
#     for y in range(len(storage_item)):
#         item_dict.update({item_dict_keys[y]: storage_item[y]})
#     storage_list.append(item_dict)
# print(storage_list)
# storage_list = list()
# increment_size = 6
# for x in range(0, len(storage_capacity)):
#     item_dict = dict()
#     storage_item = list(filter(None, storage_capacity[x].split()))
#     for y in range(len(storage_item)):
#         item_dict.update({item_dict_keys[y]: storage_item[y]})
#     storage_list.append(item_dict)
# print(storage_list)

# f = open('readme_self.excel_collection_value.txt', 'r')
# data = f.read()
# data = json.loads(data.replace("\'", "\""))
# # print(type(data))
# list_temp = list()
# for each in data :
#     temp = dict()
#     for key, value in each.items():
#         if type(value) == list:
#             strings = ''
#             for item in value:
#                 if type(item)== list:
#                     strings +=','.join(item)
#                     strings += "\r\n"
#                     value = strings
#                 elif type(item) == dict:
#                     strings +=','.join('='.join((key,val)) for (key,val) in item.items())
#                     strings += "\r\n"
#                     value = strings
#         else:
#             pass
#         # print(key, value)
#         temp.update({key: value})
#     list_temp.append(temp)

# print (list_temp)


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
