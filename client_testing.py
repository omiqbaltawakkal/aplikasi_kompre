from tkinter import *   # use lower case tkinter for python3
from threading import Thread
from time import sleep, time
import socket
import xlsxwriter
import json
import subprocess
import re
import base64
import pika

image_filepath_list = list()
report_collection = dict()

def saveImagePath():
    # f_types = [('Jpg Files', '*.jpg'),
    #            ('PNG Files', '*.png')]
    # filename = fd.askopenfilename(multiple=False)
    # filename = open('taskmanager.jpg', "rb")
    image_filepath_list.append('[FILENAME]')


def upload_file(collection_key, image_file_path):
    file_path = image_file_path
    f = open(file_path, "rb")
    im_bytes = f.read()
    im_b64 = base64.urlsafe_b64encode(im_bytes)
    report_collection.update({collection_key: im_b64})

saveImagePath()
for index, each in enumerate(image_filepath_list):
    exte = each.split(".")[1]
    upload_file("image "+str(index)+"."+exte, each)
    

with open('readme test.txt', 'w') as f:
        f.write(str(report_collection))
        f.close()


report_collection.update({"name":"iqbal"})

server_ip = "10.233.79.249"
server_port = 137


socketObject = socket.socket()
socketObject.connect((server_ip, server_port))

HTTPMessage = report_collection
bytes = str.encode(HTTPMessage)

socketObject.sendall(bytes)

while(True):
    data = socketObject.recv(1024)
    print(data)
    print("Connection closed")
    break

socketObject.close()
