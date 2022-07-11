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
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog as fd
import socket
import json
import base64
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
    brand_name = callPowershellFunc("Get-CimInstance Win32_ComputerSystem | select @{name='Merk';e={$_.Manufacturer+\" \"+$_.Model}} | ft -HideTableHeaders")
    report_collection.update({
        collection_key: returnValueParser(brand_name)
    })
    # # print("Success Getting Devices Brand Information..")


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


def getProcessorCapacity(collection_key):
    print("Getting Processor Information..")
    processor_name = callPowershellFunc("(Get-WmiObject Win32_Processor).Name")
    processor_average_utilization = callPowershellFunc(
        "(Get-WmiObject -Class win32_processor -ErrorAction Stop | Measure-Object -Property LoadPercentage -Average | Select-Object Average).Average")
    report_collection.update({
        collection_key: returnValueParser(processor_name) + ' / ' + returnValueParser(processor_average_utilization) + " %"
    })
    print("Success Getting Processor Information..")


def getMemoryCapacity(collection_key):
    print("Getting Memory Information..")
    memory_capacity = callPowershellFunc(
        "(Get-WmiObject Win32_PhysicalMemory | Measure-Object -Property capacity -Sum).sum / 1gb")

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
        collection_key: returnValueParser(memory_capacity) + " GB / "+ str(int(utilization_value)) + " %"
    })
    print("Success Getting Memory Information..")


    
def getDiskCapacity(collection_key):
    print("Getting Storage Disk Information..")
    # ft -HideTableHeaders
    storage_capacity = callPowershellFunc(
        "Get-WmiObject win32_logicaldisk | Format-Table @{label=\"Name\";e={$_.DeviceId} }, @{n=\"Size\";e={\'{0}GB\' -f [math]::Round($_.Size/1GB,2)}},@{n=\"FreeSpace\";e={\'{0}GB\' -f [math]::Round($_.FreeSpace/1GB,2)}} |  ft -HideTableHeaders")
    storage_capacity = list(
        filter(None, storage_capacity.decode("utf-8").strip().split("\r\n")))
    del storage_capacity[1]  # delete ---- ----- ---- ----
    # [DeviceId, DriveType, Size, FreeSpace ]
    item_dict_keys = storage_capacity[0].split()
    del storage_capacity[0]  # delete header
    storage_list = list()
    increment_size = 5
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
        antivirus_dict.update({"Name": " ".join(item_split[0:2])})
        antivirus_dict.update({"Last Update": " ".join(item_split[2:-1])})
        # antivirus_dict.update({"Status": item_split[-1]})
        list_antivirus_info.append(antivirus_dict)
    report_collection.update({
        collection_key: list_antivirus_info
    })
    print("Success Getting Antiviruses Information..")


def getRemoteDesktopPortStatus(collection_key):
    print("Getting Remote Desktop Information..")
    remote_desktop_port_status = callPowershellFunc(
        "if ((Get-ItemProperty \"hklm:\System\CurrentControlSet\Control\Terminal Server\").fDenyTSConnections -eq 0) { write-host \"Ya\" } else { write-host \"Tidak\" }")
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
        result = result.decode("utf-8").strip().split("\r\n")
        report_collection.update(
            {collection_key: result[0].split()[0] + " / "+ result[0].split()[1]})
    else:
        report_collection.update({collection_key: "IP Not Found"})
    # Label(second_frame, text= "Success Getting IP Address..").pack() 


def getNetworkTimeProtocol(collection_key):
    # Label(second_frame, text= "Getting Network Time Protocol..").pack()
    result = callPowershellFunc(
        "w32tm /query /status | select-string \"Source:\"")
    result = result.decode("utf-8").strip()
    ntp_status = ""
    if result :
        result = result.split(":")
        if str(result[0]) == "Source":
            ntp_status = result[1].split(",")[0]
        else:
            ntp_status = "Error - Please check manually using PowerShell and type w32tm /query /status | select-string \"Source:\""
    else:
        ntp_status = "Error - Please check manually using PowerShell and type : w32tm /query /status | select-string \"Source:\""
    report_collection.update(
        {collection_key: ntp_status.strip()})
    # print("Success Getting Network Time Protocol..")


def getScreenSaverStatus(collection_key):
    # print("Getting Screen Saver Status..")
    screen_saver_lists = callPowershellFunc(
    "Get-WmiObject -Class Win32_Desktop | Select-Object Name, ScreenSaverActive, ScreenSaverTimeout | where-object ScreenSaverActive | ft -HideTableHeader")
    screen_saver_lists = list(
        filter(None, screen_saver_lists.decode("utf-8").strip().split(" ")))
    screen_dict = dict()
    if screen_saver_lists:
        screen_dict.update({"Name": screen_saver_lists[0], "Status": "Active", "ScreenSaverTimeout": screen_saver_lists[2]})
    else :
        screen_dict.update({"Status": "Inactive"})
    report_collection.update(
        {collection_key: screen_dict})
    # print("Success Getting Screen Saver Status..")
    
def getUsbHardeningStatus(collection_key):
    hardening_status = callPowershellFunc("Get-ItemPropertyValue \"HKLM:\\SYSTEM\\CurrentControlSet\\services\\USBSTOR\" -Name \"Start\"")
    if int(hardening_status) == 3:
        report_collection.update({collection_key : "Tidak"})
    else: 
        report_collection.update({collection_key : "Ya"})
        
def getEnvironmentName(collection_key):
    environment_name = callPowershellFunc("$env:computername")
    report_collection.update({collection_key : environment_name.decode("utf-8").strip()})

def getInstalledApplication(collection_key):
    result_x86 = callPowershellFunc(
            "Get-ItemProperty HKLM:\\Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\* |  Select-Object DisplayName | ?{ $_.DisplayName -ne $null } | sort-object -Property DisplayName -Unique | Format-Table -HideTableHeaders")
    result_x87 = callPowershellFunc("Get-ItemProperty HKLM:\\Software\Wow6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall\* | Select-Object DisplayName | ?{ $_.DisplayName -ne $null } | sort-object -Property DisplayName -Unique | Format-Table -HideTableHeaders")
    result_x86 = list(filter(None, result_x86.decode("utf-8").split("\r\n")))
    result_x87 = list(filter(None, result_x87.decode("utf-8").split("\r\n")))

    installed_app = list()
    for each in result_x86 + result_x87:
        installed_app.append(each.strip())
    installed_app.sort()
    report_collection.update({ collection_key :installed_app})


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
    # print("Start Sending to Receiver..")
    bytess = str.encode(message)
    socketObject.sendall(bytess)

    while(True):
        data = socketObject.recv(1024)
        # print(data)
        print("Connection closed")
        break

    socketObject.close()

def reportDataListFunction():
    recordUpdate(report_collection, "tanggal",  str(datetime.now().strftime("%d/%m/%Y %H:%M:%S")))
    recordUpdate(report_collection, "nama", worker_name_entry.get())
    recordUpdate(report_collection, "pn", personal_number_entry.get())
    recordUpdate(report_collection, "jabatan", worker_role_entry.get())
    recordUpdate(report_collection, "kode_uker",
                worker_branch_code_entry.get().upper())
    recordUpdate(report_collection, "kondisi", clicked_cond.get())
    recordUpdate(report_collection, "bribox", clicked_bribox.get())
    getDevicesBrand("merk")
    getEnvironmentName("nama_pc")
    getOSValues("os")
    getProcessorCapacity("processor_details")
    getMemoryCapacity("ram")
    getRemoteDesktopPortStatus("rdp")
    getUsbHardeningStatus("hardening")
    getNetworkTimeProtocol("ntp")
    getLANIPAddress("ip_address")
    getScreenSaverStatus("screensaver")
    getDiskCapacity("disk")
    getAntivirusProduct("antivirus")
    getScreenSaverStatus("screensaver")


def getStartProcedures():
    server_ip = str(ip_entry.get().split(":")[0])
    server_port = int(ip_entry.get().split(":")[1])
    
    # threads = threading.Thread(target=).start()
    reportDataListFunction()
    socketObject = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    socketObject.connect((server_ip, server_port))
    # # Label(second_frame, text='Connected to designated IP for sending..').pack()
    sendDataToHost(socketObject, str(report_collection))
    message_box = tk.messagebox.showinfo("Info", "Selesai...")
    root.destroy()


# now we are required to tell Python
# for 'Main' function existence
if __name__ == '__main__':

    root = tk.Tk()
    root.title('Client-side PC Information Getter')
    # get the screen dimension
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()

    # find the center point
    center_x = int(0.45 *screen_width)
    center_y = int(0.2 * screen_height)

    # set the position of the window to the center of the screen
    #            width x height
    root.geometry(f'400x700+{center_x}+{center_y}')
    
    x =0
    ip_label = tk.Label(root, text='IP/Port Receiver     :')
    # name_label.pack(side = LEFT, fill=X, expand=True)
    ip_label.grid(column = 0, row = x,sticky = 'w', pady=20)
    

    ip_str_var = tk.StringVar()
    ip_entry = tk.Entry(root, textvariable=ip_str_var)
    # worker_name_entry.pack(side = RIGHT, fill=X, expand=True)
    ip_entry.grid(column = 1, row = x,sticky='nesw', pady=20)
    ip_entry.focus()

    name_label = tk.Label(root, text='Nama Pekerja     :')
    # name_label.pack(side = LEFT, fill=X, expand=True)
    name_label.grid(column = 0, row = x+1+0,sticky = 'w', pady=20)

    worker_name = tk.StringVar()
    worker_name_entry = tk.Entry(root, textvariable=worker_name)
    # worker_name_entry.pack(side = RIGHT, fill=X, expand=True)
    worker_name_entry.grid(column = 1, row = x+1+0,sticky='nesw', pady=20)

    personal_number_label = tk.Label(root, text='PN Pekerja        :')
    personal_number_label.grid(column = 0, row = x+1+1,sticky = 'w', pady=20)
    # personal_number_label.pack(side = LEFT,fill=X, expand=True)

    personal_number = tk.IntVar()
    personal_number_entry = tk.Entry(root, textvariable=personal_number)
    personal_number_entry.grid(column = 1, row = x+1+1, sticky='nesw', pady=20)
    # personal_number_entry.pack(side = RIGHT,fill=X, expand=True)

    worker_role_label = tk.Label(root, text='Jabatan           :')
    worker_role_label.grid(column = 0, row = x+1+2,sticky = 'w', pady=20)
    # worker_role_label.pack(side = LEFT,fill=X, expand=True)

    worker_role = tk.StringVar()
    worker_role_entry = tk.Entry(root, textvariable=worker_role)
    worker_role_entry.grid(column = 1, row = x+1+2,sticky='nesw', pady=20)
    # worker_role_entry.pack(side = RIGHT,fill=X, expand=True)

    worker_branch_code_label = tk.Label(root, text='Kode Uker          :')
    worker_branch_code_label.grid(column = 0, row = x+1+3,sticky = 'w', pady=20)
    # worker_branch_code_label.pack(side = LEFT,fill=X, expand=True)

    worker_branch_code = tk.IntVar()
    worker_branch_code_entry = tk.Entry(
        root, textvariable=worker_branch_code)
    worker_branch_code_entry.grid(column = 1, row = x+1+3,sticky='nesw', pady=20)
    # worker_branch_code_entry.pack(side = RIGHT,fill=X, expand=True)

    # open image
    
    rows = x+1+4
    condition_label = tk.Label(root, text='Kondisi Fisik    :')
    condition_label.grid(column = 0, row = rows,sticky = 'w', pady=20)
    
    cond_options = ["Baik", "Bermasalah"]
    clicked_cond = tk.StringVar()
    clicked_cond.set( "Select" )
    tk.OptionMenu( root , clicked_cond , *cond_options ).grid(column = 1, row = rows)
    rows+=1
    
    
    sticker_label = tk.Label(root, text='Ada Sticker Bribox          :')
    sticker_label.grid(column = 0, row = rows,sticky = 'w', pady=20)
    
    sticker_options = ["Ya", "Tidak"]
    clicked_bribox = tk.StringVar()
    clicked_bribox.set( "Select" )
    tk.OptionMenu( root , clicked_bribox , *sticker_options ).grid(column = 1, row = rows)
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

    ttk.Button(
        root,
        text="Submit",
        command=getStartProcedures,
    ).grid(column = 2, row = rows,sticky='nesw', pady=20)
    # start_button.pack(
    #     ipadx=5,
    #     ipady=5,
    #     expand=True
    # )

    root.mainloop()
