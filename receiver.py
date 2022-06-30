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
from tkinter.scrolledtext import ScrolledText
from threading import Thread
import traceback

from time import sleep,time


break_condition = True
heading = 'KKPA PC Pekerja'


window_width = 500
window_height = 200


class GUI():
    def __init__(self, root):
        self.running = 0    #not listening
        self.addr = None
        self.conn = None
        self.serverSocket = None
        
        self.server_ip = "10.233.79.249"
        self.server_port = 50392
        
        self.collection_data_value_only = list()
        self.header_list = list()
        
        self.screen_width = root.winfo_screenwidth()
        self.screen_height = root.winfo_screenheight()
        # find the center point
        self.center_x = int(0.2 * self.screen_width)
        self.center_y = int(0.2 * self.screen_height)

        # set the position of the window to the center of the screen
        #            width x height
        root.geometry(f'400x700+{self.center_x}+{self.center_y}')
        
        self.frame = Frame(root)
        self.frame.pack(side = LEFT, anchor = N)
        
        self.ip_port = Label(self.frame, text = "IP/PORT : ")
        self.ip_port.pack(side = LEFT, anchor = SW)
        
        self.ip_str_var = StringVar()
        self.ip_port_entry = Entry(self.frame,textvariable=self.ip_str_var)
        self.ip_port_entry.pack(side = LEFT, anchor = N)

        self.startb = Button(self.frame, text = "Start", command = self.startc)
        self.startb.pack(side = LEFT, anchor = N)

        self.stopb = Button(self.frame, text = "Stop", command = self.stopc)
        self.stopb.pack(side = LEFT, anchor = N)

        self.connectionl = Label(self.frame, text = "Stopped")
        self.connectionl.pack(side = LEFT, anchor = SW)

    def generateKKPAExcel(self):
        workbook = xlsxwriter.Workbook('kkpa.xlsx')
        worksheet = workbook.add_worksheet()
        
        for x in range(len(header_list)):
            worksheet.write(0, x, str(header_list[x]))
        
        for x in range(len(collection_data_value_only)):
            for y in range(len(collection_data_value_only[0])):
                worksheet.write(1+x, y, str(collection_data_value_only[x][y]))
            
        workbook.close()

    def generateKKPAFromRawData(self, client_raw_data):
        dict_raw_data = json.loads(client_raw_data)
        document = Document()
        owner_name = dict_raw_data.get('nama_pekerja')
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
                document.add_heading("{}".format(key.replace("_", " ").title()), 1)
                if str(key).__contains__("storage"):
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
                            document.add_paragraph("Desktop Account Name of {}, have screen save with timeout {} seconds".format(
                                item.get('Name'), item.get('ScreenSaverTimeout')), style='List Number')
                elif str(key).__contains__("ip_addre"):
                    for item in value:
                        document.add_paragraph("{} ".format(item), style='List Number')
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
        print("thread started..")
        self.running = 1
        self.server_ip = str(self.ip_port_entry.get().split(":")[0])
        self.server_port = int(self.ip_port_entry.get().split(":")[1])
        
        self.serverSocket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)

        self.serverSocket.bind((self.server_ip, self.server_port))
        self.serverSocket.listen()
        # self.serverSocket.settimeout(5)
        self.conn=None
        while self.running != 0:
            print("running ", self.running)
            if self.conn is None:
                try:
                    (self.conn, self.addr) = self.serverSocket.accept()
                    print("client is at", self.addr[0], "on port", self.addr[1])
                    self.connectionl.configure(text="CONNECTED!")
                except Exception as E:
                    print("Connect exception: "+str(E) )
                    self.conn == None
            print("self.conn " , self.conn)
            if self.conn != None:
                print ("connected to "+str(self.conn)+","+str(self.addr))
                # self.conn.settimeout(5)
                self.rc = ""
                connect_start = time() # actually, I use this for a timeout timer
                while self.rc != "done":
                    self.rc=''
                    try:
                        self.rc = self.conn.recv(1024000000).decode('utf-8')

                    except Exception as E:
                        # we can wait on the line if desired
                        print ("socket error: "+str(E))
                        self.conn = None
                    if len(self.rc):
                        # print("got data", self.rc)
                        #disini kasih generate
                        self.generateKKPAFromRawData(self.rc.replace("\'", "\""))
                        self.conn.send(str.encode("Received !!!"))
                        connect_start=time()  # reset timeout time
                    elif (self.running==0) or (time()-connect_start > 1800):
                        print ("Tired of waiting on connection!")
                        self.rc = "done"

                print ("closing connection")
                self.connectionl.configure(text="not connected.")
                self.conn.close()
                self.conn=None
                print ("connection closed.")

        print ("closing listener...")
        # self running became 0
        self.serverSocket.close()

    def startc(self):
        if self.running == 0:
            print ("Starting thread")
            self.running = 1
            self.thread=Thread(target=self.socket_thread)
            self.thread.start()
        else:
            print ("thread already started.")

    def stopc(self):
        if self.running:
            print ("stopping thread...")
            self.running = 0
            self.thread.join()
        else:
            print ("thread not running")

root = Tk()
gui = GUI(root)
root.mainloop()