# GUI
# 1. nama pekerja
# 2. jabatan
# 3. PN
# 4. Uker

# spek
# 1. merek ->> (Get-WmiObject -Class:Win32_ComputerSystem).Name + (Get-WmiObject -Class:Win32_ComputerSystem).Model
# 2. OS ->>  (Get-WMIObject win32_operatingsystem).name + (Get-WMIObject win32_operatingsystem).version + (Get-WMIObject win32_operatingsystem).OSArchitecture
# 3. processor / utilisasi ->> (Get-WmiObject Win32_Processor).Name / (Get-WmiObject -ComputerName $env:computername -Class win32_processor -ErrorAction Stop | Measure-Object -Property LoadPercentage -Average | Select-Object Average).Average
# 4. RAM /utilisasi  (Get-WmiObject Win32_PhysicalMemory | Measure-Object -Property capacity -Sum).sum / 1mb /
# 5. ALL DISK - Free space / total Get-Volume  Get-WmiObject Win32_LogicalDisk
# 6. antivirus  Get-WmiObject -Namespace root/SecurityCenter2 -ClassName AntivirusProduct | Select-Object -Property "display*", "time*", "productStat*"
# 7. saved password
# 8. remote desktop connection if ((Get-ItemProperty "hklm:\System\CurrentControlSet\Control\Terminal Server").fDenyTSConnections -eq 0) { write-host "RDP is Enabled" } else { write-host "RDP is NOT enabled" }
# 9. IP   ->> (Get-NetIPAddress -AddressFamily IPV4 -InterfaceAlias "Ethernet*").IPAddress
# 10. clear screen and screen saver Get-WmiObject -Class Win32_Desktop
# 11. NTP w32tm /query /computer:$env:computername /status
# Get-WmiObject -List

import subprocess
from tkinter import *
from tkinter import ttk
from tkinter import filedialog as fd
from tkinter.messagebox import showinfo
import socket
import json
from PIL import Image, ImageTk
import base64
from pathlib import Path
import asyncio
from datetime import datetime


window_width = 200
window_height = 400

server_ip = "127.0.0.1"
server_port = 1002

# report variables
report_collection = dict()

image_filepath_list = list()

image_dict_list = ["Task Manager", "fisik", "bribox"]
# "stiker bribox", "dxdiag image", "saved password"


def recordUpdate(dictionary, collection_key, collection_value):
    dictionary.update({collection_key: collection_value})


def returnValueParser(return_value):
    return return_value.decode("utf-8").strip().replace("\r\n", "")


def callPowershellFunc(cmd):
    result = subprocess.Popen(
        ["powershell", "-Command", cmd], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    out, err = result.communicate()
    return out


def getDevicesBrand(collection_key):
    print("Getting Devices Brand Information..")
    brand_name = callPowershellFunc(
        "(Get-WmiObject -Class:Win32_ComputerSystem).Name")
    brand_model = callPowershellFunc(
        "(Get-WmiObject -Class:Win32_ComputerSystem).Model")
    report_collection.update({
        collection_key: returnValueParser(
            brand_name) + " " + returnValueParser(brand_model)
    })
    print("Success Getting Devices Brand Information..")


def getOSValues(collection_key):
    print("Getting Operating System Information..")
    # Microsoft Windows 10 Home Single Language|C:\Windows|\Device\Harddisk0\Partition3
    os_name = callPowershellFunc("(Get-WMIObject win32_operatingsystem).name")
    os_version = callPowershellFunc(
        "(Get-WMIObject win32_operatingsystem).version")
    os_architecture = callPowershellFunc(
        "(Get-WMIObject win32_operatingsystem).OSArchitecture")
    report_collection.update({
        collection_key: returnValueParser(os_name).split(
            "|")[0] + " " + returnValueParser(os_version) + " " + returnValueParser(os_architecture)
    })
    print("Success Getting Operating System Information..")


def getProcessorCapacity(collection_key, collection_key_2):
    print("Getting Processor Information..")
    processor_name = callPowershellFunc("(Get-WmiObject Win32_Processor).Name")
    processor_average_utilization = callPowershellFunc(
        "(Get-WmiObject -Class win32_processor -ErrorAction Stop | Measure-Object -Property LoadPercentage -Average | Select-Object Average).Average")
    report_collection.update({
        collection_key: returnValueParser(processor_name),
        collection_key_2: returnValueParser(
            processor_average_utilization) + "%"
    })
    print("Success Getting Processor Information..")


def getMemoryCapacity(collection_key, collection_key_2):
    print("Getting Memory Information..")
    memory_capacity = callPowershellFunc(
        "(Get-WmiObject Win32_PhysicalMemory | Measure-Object -Property capacity -Sum).sum / 1mb")

    total_visible_memory_size = callPowershellFunc(
        "(Get-WmiObject -Class WIN32_OperatingSystem).TotalVisibleMemorySize")
    total_visible_memory_size = float(
        total_visible_memory_size.decode("utf-8"))
    free_physical_memory_size = callPowershellFunc(
        "(Get-WmiObject -Class WIN32_OperatingSystem).FreePhysicalMemory")
    free_physical_memory_size = float(
        free_physical_memory_size.decode("utf-8"))
    utilization_value = (total_visible_memory_size -
                         free_physical_memory_size)*100 / total_visible_memory_size
    report_collection.update({
        collection_key: returnValueParser(memory_capacity) + " MB",
        collection_key_2: str(int(utilization_value)) + "%"
    })
    print("Success Getting Memory Information..")


def getDiskCapacity(collection_key):
    print("Getting Storage Disk Information..")
    storage_capacity = callPowershellFunc(
        "Get-WmiObject win32_logicaldisk | Format-Table DeviceId, DriveType, @{n=\"Size\";e={[math]::Round($_.Size/1GB,2)}},@{n=\"FreeSpace\";e={[math]::Round($_.FreeSpace/1GB,2)}}")
    storage_capacity = list(
        filter(None, storage_capacity.decode("utf-8").strip().split("\r\n")))
    del storage_capacity[1]  # delete ---- ----- ---- ----
    # [DeviceId, DriveType, Size, FreeSpace ]
    item_dict_keys = storage_capacity[0].split()
    del storage_capacity[0]  # delete header
    storage_list = list()
    increment_size = 6
    for x in range(0, len(storage_capacity)):
        item_dict = dict()
        storage_item = list(filter(None, storage_capacity[x].split()))
        for y in range(len(storage_item)):
            item_dict.update({item_dict_keys[y] : storage_item[y]})
        storage_list.append(item_dict)
    report_collection.update({
        collection_key: storage_list
    })
    print("Success Getting Storage Disk Information..")


def getAntivirusProduct(collection_key):
    print("Getting Antiviruses Information..")
    list_of_antiviruses = callPowershellFunc(
        "Get-WmiObject -Namespace root/SecurityCenter2 -ClassName AntivirusProduct | Select-Object -Property display*, time*, productStat* | ft -HideTableHeaders")
    list_of_antiviruses = list_of_antiviruses.decode(
        "utf-8").strip().split("\r\n")  # wajib
    list_antivirus_info = list()
    for item in list_of_antiviruses:
        antivirus_dict = dict()
        item_split = item.split()
        antivirus_dict.update({"AntivirusName": " ".join(item_split[0:2])})
        antivirus_dict.update({"LastUpdate": " ".join(item_split[2:-1])})
        antivirus_dict.update({"Status": item_split[-1]})
        list_antivirus_info.append(antivirus_dict)
    report_collection.update({
        collection_key: list_antivirus_info
    })
    print("Success Getting Antiviruses Information..")


def getRemoteDesktopPortStatus(collection_key):
    print("Getting Remote Desktop Information..")
    remote_desktop_port_status = callPowershellFunc(
        "if ((Get-ItemProperty \"hklm:\System\CurrentControlSet\Control\Terminal Server\").fDenyTSConnections -eq 0) { write-host \"Remote Desktop Port is Enabled\" } else { write-host \"Remote Desktop Port is disabled\" }")
    report_collection.update({
        collection_key: returnValueParser(remote_desktop_port_status)
    })
    print("Success Getting Remote Desktop Information..")


def getLANIPAddress(collection_key):
    # Label(second_frame, text= "Getting IP Address..").pack() 
    result = callPowershellFunc(
        "Get-WmiObject win32_networkadapterconfiguration | Select-Object -Property @{Name = 'IPAddress' ; Expression = {($PSItem.IPAddress[0])}}, MacAddress | Where IPAddress -NE $null | ft -HideTableHeaders")
    #Get-WmiObject Win32_NetworkAdapterConfiguration -Filter "DHCPEnabled=$True" | Where-Object {$_.IPEnabled -AND $_.IPAddress -gt 0} |Select-object IPAddress, MACAddress
    if result:
        report_collection.update(
            {collection_key: result.decode("utf-8").strip().split(" ")})
    else:
        report_collection.update({collection_key: "IP Not Found"})
    # Label(second_frame, text= "Success Getting IP Address..").pack() 


def getNetworkTimeProtocol(collection_key):
    # Label(second_frame, text= "Getting Network Time Protocol..").pack()
    result = callPowershellFunc(
        "w32tm /query /computer:$env:computername /status")
    result = result.decode("utf-8").strip().split("\r\n")[0]
    result = list(filter(None, result.split("\n")))
    ntp_dict = dict()
    ntp_status = "Synced"
    for item in result:
        item_split = item.split(":")
        if str(item_split[0]) == "Source":
            ntp_status = ntp_status + " from " + "".join(item_split[1:]).strip() + ""
        if str(item_split[0]) == "Last Successful Sync Time":
            ntp_status = ntp_status + " at " + " : " .join(item_split[1:]).strip()
    report_collection.update(
        {collection_key: ntp_status})
    print("Success Getting Network Time Protocol..")


def getScreenSaverStatus(collection_key):
    print("Getting Screen Saver Status..")
    screen_saver_lists = callPowershellFunc(
        "Get-WmiObject -Class Win32_Desktop")
    screen_saver_lists = list(
        filter(None, screen_saver_lists.decode("utf-8").strip().split("\r\n")))
    screen_saver = list()
    increment_size = 5
    for x in range(0, len(screen_saver_lists), increment_size):
        item_dict = dict()
        for item in screen_saver_lists[x:x+increment_size]:
            item_split = item.split(":")
            item_dict.update({item_split[0].strip(): item_split[1].strip()})
        screen_saver.append(item_dict)
    report_collection.update(
        {collection_key: screen_saver})
    print("Success Getting Screen Saver Status..")
    
def getUsbHardeningStatus(collection_key):
    hardening_status = callPowershellFunc("Get-ItemPropertyValue \"HKLM:\\SYSTEM\\CurrentControlSet\\services\\USBSTOR\" -Name \"Start\"")
    if int(hardening_status) == 3:
        report_collection.update({collection_key : "Enabled"})
    else: 
        report_collection.update({collection_key : "Disabled"})


def saveImagePath():
    f_types = [('Jpg Files', '*.jpg'),
               ('PNG Files', '*.png')]
    filename = fd.askopenfilename(multiple=False)
    image_filepath_list.append(filename)


def upload_file(collection_key, image_file_path):
    file_path = image_file_path
    f = open(file_path, "rb")
    im_bytes = f.read()
    im_b64 = base64.b64encode(im_bytes).decode("utf8")
    report_collection.update({collection_key: im_b64})


def sendDataToHost(socketObject, message):
    print("Start Sending to Receiver..")
    bytess = str.encode(message)
    print(len(bytess))
    socketObject.sendall(bytess)

    while(True):
        data = socketObject.recv(1024)
        print(data)
        print("Connection closed")
        break

    socketObject.close()

def powershellListFunction():
    getLANIPAddress("ip_address")    
    getDevicesBrand("computer_name")
    getOSValues("os_values")
    getProcessorCapacity("processor_capacity", "processor_utilization")
    getMemoryCapacity("memory_capacity", "memory_utilization")
    getRemoteDesktopPortStatus("remote_desktop_status")
    getNetworkTimeProtocol("network_time_protocol_Status")
    getDiskCapacity("list_of_storage_information")
    getAntivirusProduct("list_of_antiviruses")
    getScreenSaverStatus("list_of_screen_saver_status")


def getPersonalInformation():
    server_ip = str(ip_entry.get().split(":")[0])
    server_port = int(ip_entry.get().split(":")[1])
    # Label(second_frame, text='Start Gather').pack()
    # if not report_collection:
    recordUpdate(report_collection, "nama_pekerja", worker_name_entry.get())
    recordUpdate(report_collection, "tanggal",  str(datetime.now().strftime("%d/%m/%Y %H:%M:%S")))
    recordUpdate(report_collection, "pN_pekerja", personal_number_entry.get())
    recordUpdate(report_collection, "jabatan_pekerja", worker_role_entry.get())
    recordUpdate(report_collection, "kode_uker",
                worker_branch_code_entry.get())
    recordUpdate(report_collection, "condition", clicked_cond.get())
    recordUpdate(report_collection, "bribox_sticker", clicked_bribox.get())
    powershellListFunction()
    # for x in range(len(image_filepath_list)):
    #     collection_name = (
    #         "image_" + "_".join(image_dict_list[x].split(" "))).lower()
    #     upload_file(collection_name, image_filepath_list[x])

    with open('readme '+worker_name_entry.get()+'.txt', 'w') as f:
        f.write(str(report_collection))
        f.close()
    socketObject = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    socketObject.connect((server_ip, server_port))
    # # Label(second_frame, text='Connected to designated IP for sending..').pack()
    # sendDataToHost(socketObject, str(report_collection))

# now we are required to tell Python
# for 'Main' function existence
if __name__ == '__main__':

    root = Tk()
    root.title(string='LHPK')
    
    # get the screen dimension
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()

    # find the center point
    center_x = int(0.45 *screen_width)
    center_y = int(0.2 * screen_height)

    # set the position of the window to the center of the screen
    #            width x height
    root.geometry(f'400x700+{center_x}+{center_y}')
    
    # main_frame = Frame(root)
    # main_frame.grid(row=0, column=0, sticky="news")

    # my_canvas = Canvas(main_frame)
    # my_canvas.grid(row=0, column=0, sticky="news")

    # my_scrollbar = ttk.Scrollbar(
    #     main_frame, orient=VERTICAL, command=my_canvas.yview)
    # my_scrollbar.grid(row=0, column=1, sticky="ns")

    # my_canvas.configure(yscrollcommand=my_scrollbar.set)
    # my_canvas.configure(scrollregion=my_canvas.bbox("all"))

    # second_frame = Frame(my_canvas)

    # my_canvas.create_window((0, 0), window=second_frame, anchor="nw")

    x =0
    
    ip_label = Label(root, text='IP/Port Receiver     :')
    # name_label.pack(side = LEFT, fill=X, expand=True)
    ip_label.grid(column = 0, row = x,sticky = 'w', pady=20)
    

    ip_str_var = StringVar()
    ip_entry = Entry(root, textvariable=ip_str_var)
    # worker_name_entry.pack(side = RIGHT, fill=X, expand=True)
    ip_entry.grid(column = 1, row = x,sticky='nesw', pady=20)
    ip_entry.focus()
    ########### 

    name_label = Label(root, text='Nama Pekerja     :')
    # name_label.pack(side = LEFT, fill=X, expand=True)
    name_label.grid(column = 0, row = x+1+0,sticky = 'w', pady=20)

    worker_name = StringVar()
    worker_name_entry = Entry(root, textvariable=worker_name)
    # worker_name_entry.pack(side = RIGHT, fill=X, expand=True)
    worker_name_entry.grid(column = 1, row = x+1+0,sticky='nesw', pady=20)

    personal_number_label = Label(root, text='PN Pekerja        :')
    personal_number_label.grid(column = 0, row = x+1+1,sticky = 'w', pady=20)
    # personal_number_label.pack(side = LEFT,fill=X, expand=True)

    personal_number = IntVar()
    personal_number_entry = Entry(root, textvariable=personal_number)
    personal_number_entry.grid(column = 1, row = x+1+1, sticky='nesw', pady=20)
    # personal_number_entry.pack(side = RIGHT,fill=X, expand=True)

    worker_role_label = Label(root, text='Jabatan           :')
    worker_role_label.grid(column = 0, row = x+1+2,sticky = 'w', pady=20)
    # worker_role_label.pack(side = LEFT,fill=X, expand=True)

    worker_role = StringVar()
    worker_role_entry = Entry(root, textvariable=worker_role)
    worker_role_entry.grid(column = 1, row = x+1+2,sticky='nesw', pady=20)
    # worker_role_entry.pack(side = RIGHT,fill=X, expand=True)

    worker_branch_code_label = Label(root, text='Kode Uker          :')
    worker_branch_code_label.grid(column = 0, row = x+1+3,sticky = 'w', pady=20)
    # worker_branch_code_label.pack(side = LEFT,fill=X, expand=True)

    worker_branch_code = IntVar()
    worker_branch_code_entry = Entry(
        root, textvariable=worker_branch_code)
    worker_branch_code_entry.grid(column = 1, row = x+1+3,sticky='nesw', pady=20)
    # worker_branch_code_entry.pack(side = RIGHT,fill=X, expand=True)

    # open image
    
    rows = x+1+4
    cond_options = ["Bermasalah", "Baik"]
    sticker_options = ["Ya", "Tidak"]
    
    clicked_cond = StringVar()
    clicked_cond.set( "Select" )
    condition_drop = OptionMenu( root , clicked_cond , *cond_options ).grid(column = 1, row = rows)
    rows+=1
    
    clicked_bribox = StringVar()
    clicked_bribox.set( "Select" )
    bribox_drop = OptionMenu( root , clicked_bribox , *sticker_options ).grid(column = 1, row = rows)
    rows+=1
    
    # for item in image_dict_list:
    #     attachment_label = Label(root, text=f'Attachment {item} : ')
    #     attachment_label.grid(column = 0, row = rows, sticky='w', pady=20)
    #     # attachment_label.pack(fill=X, expand=True)

    #     attachment_button = ttk.Button(
    #         root,
    #         text='Open Attachment Image',
    #         command=saveImagePath
    #     ).grid(column = 1, row = rows, sticky='nesw', pady=20)

    start_button = ttk.Button(
        root,
        text="Submit",
        command=getPersonalInformation,
    ).grid(column = 2, row = rows,sticky='nesw', pady=20)
    # start_button.pack(
    #     ipadx=5,
    #     ipady=5,
    #     expand=True
    # )

    root.mainloop()
