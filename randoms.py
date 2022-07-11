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

server_ip = "10.233.79.249"
server_port = 137

serverSocket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)

serverSocket.bind((server_ip, server_port))
serverSocket.listen()
# serverSocket.settimeout(5)
count = 0
while(True):
    (conn, addr) = serverSocket.accept()
    temp =''
    # try:
    while True:
        rc = conn.recv(40960)
        if not rc:
            break
        temp += rc.decode()
    conn.sendall(b'end:')
        
    with open('readme test.txt', 'w') as f:
        f.write(str(temp))
        f.close()
    
    conn.close()
    break
    
        