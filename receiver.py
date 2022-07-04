# import the socket module
import json
import socket
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Cm, Pt
from PIL import Image, ImageTk
import base64
import xlsxwriter
from tkinter import *
from tkinter import ttk
import tkinter.scrolledtext as scrolledtext
from threading import Thread
import traceback
from time import sleep, time
import pandas as pd


break_condition = True
heading = 'KKPA PC Pekerja'


window_width = 500
window_height = 200


class GUI():
    def __init__(self, root):
        self.running = 0  # not listening
        self.addr = None
        self.conn = None
        self.serverSocket = None

        self.server_ip = "10.233.79.249"
        self.server_port = 50392

        self.collection_data_value_only = list()
        self.header_list = list()
        
        self.excel_collection_value = list()

        self.screen_width = root.winfo_screenwidth()
        self.screen_height = root.winfo_screenheight()
        # find the center point
        self.center_x = int(0.2 * self.screen_width)
        self.center_y = int(0.2 * self.screen_height)

        # set the position of the window to the center of the screen
        #            width x height
        root.geometry(f'1000x200+{self.center_x}+{self.center_y}')

        self.frame = Frame(root)
        self.frame.pack(side=LEFT, anchor=N)

        self.ip_port = Label(self.frame, text="IP/PORT : ")
        self.ip_port.pack(side=LEFT, anchor=SW)

        self.ip_str_var = StringVar()
        self.ip_port_entry = Entry(self.frame, textvariable=self.ip_str_var)
        self.ip_port_entry.pack(side=LEFT, anchor=N)
        
        self.max_number = Label(self.frame, text="Max Size : ")
        self.max_number.pack(side=LEFT, anchor=SW)

        self.max_number_var = IntVar()
        self.max_number_entry = Entry(self.frame, textvariable=self.max_number_var)
        self.max_number_entry.pack(side=LEFT, anchor=N)
        
        self.startb = Button(self.frame, text="Start", command=self.startc)
        self.startb.pack(side=LEFT, anchor=N)

        self.generate = Button(self.frame, text="Generate", command=self.generateKKPAExcel)
        self.generate.pack(side=LEFT, anchor=N)

        # self.connectionl = Label(self.frame, text="Not Started")
        # self.connectionl.pack(side=LEFT, anchor=SW)
        self.textboxes = scrolledtext.ScrolledText(root, undo=True)
        self.textboxes.pack(expand=True, fill='both')
        self.addToTextbox("Not Started")
    
    def addToTextbox(self, message):
        self.textboxes.config(state=NORMAL)
        self.textboxes.insert(END, message +"\n")
        self.textboxes.config(state=DISABLED)
        
    def generateKKPAExcel(self):
        data = None
        if (self.excel_collection_value):
            data = self.excel_collection_value
        else: 
            f = open('readme_self.excel_collection_value.txt', 'r')
            data = f.read()
            data = json.loads(data.replace("\'", "\""))
            
        #asumsi end process = 1 uker -> data[0].get("kode_uker") == data[N].get("kode_uker")
        kode_uker = data[0].get("kode_uker")
        
        list_temp = list()
        for each in data :
            temp = dict()
            for key, value in each.items():
                strings = ''
                for item in value:
                    if type(item)== list:
                        strings +=','.join(item)
                        strings += "\n"
                        value = strings
                    elif type(item) == dict:
                        strings +="\r\n".join(':'.join((key,val)) for (key,val) in item.items())
                        strings += "\n\n"
                        value = strings
                else:
                    pass
                temp.update({key.replace("_"," ").title(): value})
            list_temp.append(temp)
        data = list_temp
        
        df = pd.DataFrame(data=data)
        writer = pd.ExcelWriter('kkpa.xlsx', engine='xlsxwriter')
        df.to_excel(writer, index=False, startrow=2, sheet_name='Sheet1')
        worksheet = writer.sheets['Sheet1']
        worksheet.write('A1', "Unit Kerja : " + kode_uker)
        writer.save()

    def generateKKPAFromRawData(self, client_raw_data):
        self.addToTextbox("generate KKPA start ...")
        dict_raw_data = client_raw_data
        document = Document()
        owner_name = dict_raw_data.get('nama')
        # Header
        heading_style = document.styles['Body Text']
        head = document.add_paragraph(style=heading_style).add_run(
            f'{heading}' + ' ' + owner_name)
        document.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
        head.font.size = Pt(20)
        head.font.bold = True
        image_width = Cm(10)

        for key, value in dict_raw_data.items():
            if type(value) == list:
                document.add_heading("{}".format(
                    key.replace("_", " ").title()), 1)
                if str(key).__contains__("disk"):
                    for item in value:
                        document.add_paragraph(" Disk {} Total / Free Space : {} GB / {} GB ".format(
                            item.get('DeviceId'), item.get('Size'), item.get('FreeSpace')), style='List Number')
                elif str(key).__contains__("antivirus"):
                    for item in value:
                        document.add_paragraph("{}, updated on {}".format(
                            item.get('AntivirusName'), item.get('LastUpdate')), style='List Number')
                elif str(key).__contains__("saver"):
                    for item in value:
                        if item.get('ScreenSaverTimeout'):
                            document.add_paragraph("Desktop Account Name of {}, have screensaver with timeout {} seconds".format(
                                item.get('Name'), item.get('ScreenSaverTimeout')), style='List Number')
                elif str(key).__contains__("ip_addre"):
                    for item in value:
                        document.add_paragraph(
                            "{} ({}) ".format(item[0], item[1]), style='List Number')
                else:
                    for item in value:
                        document.add_paragraph("", style='List Number')
                        for key, value in item.items():
                            document.add_paragraph(" {} : {} ".format(
                                key, str(value)), style='List Bullet 2')
            elif str(key).split("_")[0] == 'image':
                image_name = owner_name + '_' + str(key)+'.jpg'
                base64_img_bytes = value.encode('utf-8')
                with open(image_name, 'wb') as file_to_save:
                    decoded_image_data = base64.decodebytes(base64_img_bytes)
                    file_to_save.write(decoded_image_data)
                    # " ".join()
                document.add_paragraph(
                    ("Screen Capture " + " ".join(str(key).split("_")[1:])).title())
                document.add_picture(image_name, width=image_width)
            else:
                doc = document.add_paragraph("")
                doc.add_run(
                    key.replace("_", " ").title()).bold = True
                doc.add_run(" : {} ".format(value))
        document.save("kkpa " + owner_name + ".docx")

    def socket_thread(self):
        self.addToTextbox("thread started..")
        self.server_ip = str(self.ip_port_entry.get().split(":")[0])
        self.server_port = int(self.ip_port_entry.get().split(":")[1])
        self.max_number = int(self.max_number_entry.get())

        self.serverSocket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)

        self.serverSocket.bind((self.server_ip, self.server_port))
        self.serverSocket.listen()
        # self.serverSocket.settimeout(5)
        self.count = 0
        while(self.running == 1):
            (self.conn, self.addr) = self.serverSocket.accept()
            
            self.addToTextbox("CONNECTED!")
            self.count = self.count + 1
            
            self.addToTextbox("Accepted {} connections so far, from {} \n".format(self.count, self.addr[0]))
            
            # try:
            while True:
                self.rc = self.conn.recv(40960)
                if self.rc != b'':
                    self.raw_data = self.rc.decode().replace("\'", "\"")
                    self.dic_data = json.loads(self.raw_data)
                    self.generateKKPAFromRawData(self.dic_data)
                    self.excel_collection_value.append(self.dic_data)
                    self.conn.sendall(str.encode("Received !!"))
                    self.addToTextbox("Received!")
                    self.textboxes.see(END)
                    self.conn.close()
                break
            if self.count == self.max_number:
                self.running = 0
                break
        if self.excel_collection_value:
            with open('readme_self.excel_collection_value.txt', 'w') as f:
                f.write(str(self.excel_collection_value))
                f.close()
        self.addToTextbox("Ended!")
        self.textboxes.see(END)
        self.generateKKPAExcel()
        self.serverSocket.close()

    def startc(self):
        if self.running == 0:
            self.addToTextbox("Starting thread")
            self.running = 1
            self.threads = Thread(target=self.socket_thread, daemon=True)
            self.threads.start()
        else:
            self.addToTextbox("thread already started.")

    def stopc(self):
        if self.running == 1:
            sself.addToTextbox("stopping thread...")
            self.running = 0
            # self.threads.join()
        else:
            self.addToTextbox("thread not running...")


root = Tk()
gui = GUI(root)
root.mainloop()