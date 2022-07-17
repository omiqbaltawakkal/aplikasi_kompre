from tkinter import *   # use lower case tkinter for python3
from threading import Thread
from time import sleep, time
import socket
import xlsxwriter
import json
import subprocess
import re
import base64
import io
import bson
import struct



# def callPowershellFunc(cmd):
#     result = subprocess.Popen(
#         ["powershell", "-Command", cmd], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
#     out, err = result.communicate()
#     return out

# os_name = callPowershellFunc("(New-Object -com \"Microsoft.Update.AutoUpdate\").Results | select-object \"*Install*\" | ft -HideTableHeaders")

# print(os_name.decode().strip())