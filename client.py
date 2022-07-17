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
from tkinter import ttk, scrolledtext
from tkinter import filedialog as fd
import socket
import json
from base64 import b64encode
from datetime import datetime
import struct
import time


window_width = 200
window_height = 400

server_ip = "127.0.0.1"
server_port = 1002

# report variables
report_collection = dict()

image_filepath_list = list()

image_dict_list = ["Kondisi Fisik",
                   "Stiker Pengadaan dan BRIBOX", "Saved Password"
                   ]


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
        "Get-WMIObject Win32_ComputerSystem | select @{name='Merk';e={$_.Manufacturer+\" \"+$_.Model}} | ft -HideTableHeaders")
    report_collection.update({
        collection_key: returnValueParser(brand_name)
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
    last_installation_date = callPowershellFunc("(New-Object -com \"Microsoft.Update.AutoUpdate\").Results | select-object \"*Install*\" | ft -HideTableHeaders")
    temp = dict()
    temp.update({
        "Name": returnValueParser(os_name).split(
            "|")[0] + " " + returnValueParser(os_version) + " " + returnValueParser(os_architecture),
        "Last Update Installation Date": returnValueParser(last_installation_date)
    })
    report_collection.update({collection_key: temp})
    print("Success Getting Operating System Information..")


def getProcessorCapacity(collection_key):
    print("Getting Processor Information..")
    processor_name = callPowershellFunc("(Get-WmiObject Win32_Processor).Name")
    processor_average_utilization = callPowershellFunc(
        "(Get-WmiObject -Class win32_processor -ErrorAction Stop | Measure-Object -Property LoadPercentage -Average | Select-Object Average).Average")
    report_collection.update({
        collection_key: returnValueParser(
            processor_name) + ' / ' + returnValueParser(processor_average_utilization) + " %"
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
        collection_key: returnValueParser(
            memory_capacity) + " GB / " + str(int(utilization_value)) + " %"
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
            item_dict.update({item_dict_keys[y]: storage_item[y]})
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
    print("Getting IP Address..")
    result = callPowershellFunc(
        "Get-WmiObject win32_networkadapterconfiguration | Select-Object -Property @{Name = 'IPAddress' ; Expression = {($PSItem.IPAddress[0])}}, MacAddress | Where IPAddress -NE $null | ft -HideTableHeaders")
    # Get-WmiObject Win32_NetworkAdapterConfiguration -Filter "DHCPEnabled=$True" | Where-Object {$_.IPEnabled -AND $_.IPAddress -gt 0} |Select-object IPAddress, MACAddress
    if result:
        result = result.decode("utf-8").strip().split("\r\n")
        report_collection.update(
            {collection_key: result[0].split()[0] + " / " + result[0].split()[1]})
    else:
        report_collection.update({collection_key: "IP Not Found"})
    print("Success Getting IP Address..")


def getNetworkTimeProtocol(collection_key):
    print("Getting Network Time Protocol..")
    result = callPowershellFunc(
        "w32tm /query /status | select-string \"Source:\"")
    result = result.decode("utf-8").strip()
    ntp_status = ""
    if result:
        result = result.split(":")
        if str(result[0]) == "Source":
            ntp_status = result[1].split(",")[0]
        else:
            ntp_status = "Error - Please check manually using PowerShell and type w32tm /query /status | select-string \"Source:\""
    else:
        ntp_status = "Error - Please check manually using PowerShell and type : w32tm /query /status | select-string \"Source:\""
    report_collection.update(
        {collection_key: ntp_status.strip()})
    print("Success Getting Network Time Protocol..")


def getScreenSaverStatus(collection_key):
    print("Getting Screen Saver Status..")
    screen_saver_lists = callPowershellFunc(
        "Get-WmiObject -Class Win32_Desktop | Select-Object Name, ScreenSaverActive, ScreenSaverTimeout | where-object ScreenSaverActive | ft -HideTableHeader")
    screen_dict = dict()
    if screen_saver_lists:
        screen_saver_lists = list(
            filter(None, screen_saver_lists.decode("utf-8").strip().split(" ")))
        screen_dict.update(
            {"Name": screen_saver_lists[0], "Status": "Active", "ScreenSaverTimeout": screen_saver_lists[2]})
    else:
        screen_dict.update({"Status": "Inactive"})
    report_collection.update(
        {collection_key: screen_dict})
    print("Success Getting Screen Saver Status..")


def getUsbHardeningStatus(collection_key):
    print("Getting Hardening Status..")
    hardening_status = callPowershellFunc(
        "Get-ItemPropertyValue \"HKLM:\\SYSTEM\\CurrentControlSet\\services\\USBSTOR\" -Name \"Start\"")
    if int(hardening_status) == 3:
        report_collection.update({collection_key: "Tidak"})
    else:
        report_collection.update({collection_key: "Ya"})
    print("Success Getting Hardening Status..")


def getEnvironmentName(collection_key):
    environment_name = callPowershellFunc("$env:computername")
    report_collection.update(
        {collection_key: environment_name.decode("utf-8").strip()})


def getInstalledApplication(collection_key):
    print("Getting Installed Applications..")
    result_x86 = callPowershellFunc(
        "Get-ItemProperty HKLM:\\Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\* |  Select-Object DisplayName | ?{ $_.DisplayName -ne $null } | sort-object -Property DisplayName -Unique | Format-Table -HideTableHeaders")
    result_x87 = callPowershellFunc(
        "Get-ItemProperty HKLM:\\Software\Wow6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall\* | Select-Object DisplayName | ?{ $_.DisplayName -ne $null } | sort-object -Property DisplayName -Unique | Format-Table -HideTableHeaders")
    result_x86 = list(filter(None, result_x86.decode("utf-8").split("\r\n")))
    result_x87 = list(filter(None, result_x87.decode("utf-8").split("\r\n")))
    print("Success Getting Installed Applications..")

    installed_app = list()
    for each in result_x86 + result_x87:
        installed_app.append(each.strip())
    installed_app.sort()
    report_collection.update({collection_key: installed_app})


def saveImagePath():
    f_types = [('JPG Files', '*.jpg'),
               ('PNG Files', '*.png')]
    filename = fd.askopenfilename(multiple=False, filetypes=f_types)
    image_filepath_list.append(filename)


def upload_file(collection_key, image_file_path):
    f = open(image_file_path, "rb")
    im_bytes = f.read()
    base64_bytes = b64encode(im_bytes)
    my_file = base64_bytes.decode("utf-8")
    report_collection.update({collection_key: my_file})


def sendDataToHost(socketObject, message):
    print("Start Sending to Receiver..")
    file_data = json.dumps(message).encode("utf-8")
    socketObject.sendall(struct.pack(">I", len(file_data)))
    socketObject.sendall(file_data)

    while True:
        data = socketObject.recv(1024)
        print(data)
        print("Connection closed")
        break

    socketObject.close()


def reportDataListFunction():
    print("Start Function...")
    recordUpdate(report_collection, "tanggal",  str(
        datetime.now().strftime("%d/%m/%Y %H:%M:%S")))
    recordUpdate(report_collection, "nama", worker_name_entry.get())
    recordUpdate(report_collection, "pn", personal_number_entry.get())
    recordUpdate(report_collection, "jabatan", worker_role_entry.get())
    recordUpdate(report_collection, "kode_uker",
                 worker_branch_code_entry.get().upper())
    recordUpdate(report_collection, "kondisi", clicked_cond.get())
    recordUpdate(report_collection, "bribox", clicked_bribox.get())
    getDevicesBrand("merk")
    getEnvironmentName("nama_pc")
    getOSValues("operating_system")
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
    getInstalledApplication("application")
    for image_files in image_filepath_list:
        key = "image_"+(str(image_files).rsplit("/", maxsplit=1)[-1]).split(".")[0]
        upload_file(key, image_files)
    recordUpdate(report_collection, "informasi_tambahan",
                 text_area.get(0.0, tk.END))
    with open('readme '+worker_name_entry.get()+'.txt', 'w') as f:
        f.write(str(report_collection))
        f.close()
        
    print("End Function...")
        
#   perbandingan waktu - kasih timestamp -> ceklis
#   validitas - dikelabuin -> BELUM BISA?!
#   human error - salah penginputan -> gimana?
#   efisiensi penggunaan - TASK MANAGER

def getStartProcedures():
    start = time.time()
    server_ip = str(ip_entry.get().split(":")[0])
    server_port = int(ip_entry.get().split(":")[1])

    # threads = threading.Thread(target=).start()
    reportDataListFunction()
    socketObject = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    socketObject.connect((server_ip, server_port))
    # # Label(second_frame, text='Connected to designated IP for sending..').pack()
    sendDataToHost(socketObject, report_collection)
    end = time.time()
    message_box = tk.messagebox.showinfo("Info", "Selesai dalam waktu {0:.2f} detik.".format(end-start))
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
    center_x = int(0.35 * screen_width)
    center_y = int(0.1 * screen_height)

    # set the position of the window to the center of the screen
    #            width x height
    root.geometry(f'500x400+{center_x}+{center_y}')
    main_frame = tk.Frame(root)
    main_frame.pack(fill=tk.BOTH, expand=1)
    
    my_canvas = tk.Canvas(main_frame)
    my_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=1)
    
    my_scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=my_canvas.yview)
    my_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    
    my_canvas.configure(yscrollcommand=my_scrollbar.set)
    my_canvas.bind('<Configure>', lambda e: my_canvas.configure(scrollregion=my_canvas.bbox("all")))
    
    second_frame = tk.Frame(my_canvas)
    my_canvas.create_window((0,0), window=second_frame, anchor = 'nw')
        
    x = 0

    ip_label = tk.Label(second_frame, text='IP/Port Receiver  :')
    # ip_label.pack(side = 'left', fill='x', expand=True)
    ip_label.grid(column=0, row=x, sticky='w', pady=20)
    
    ip_str_var = tk.StringVar()
    ip_entry = tk.Entry(second_frame, textvariable=ip_str_var)
    # # worker_name_entry.pack(side = RIGHT, fill=X, expand=True)
    ip_entry.grid(column=1, row=x, sticky='nesw', pady=20)
    ip_entry.focus()

    name_label = tk.Label(second_frame, text='Nama Pekerja  :')
    # name_label.pack(side = LEFT, fill=X, expand=True)
    name_label.grid(column=0, row=x+1+0, sticky='w', pady=20)

    worker_name = tk.StringVar()
    worker_name_entry = tk.Entry(second_frame, textvariable=worker_name)
    # worker_name_entry.pack(side = RIGHT, fill=X, expand=True)
    worker_name_entry.grid(column=1, row=x+1+0, sticky='nesw', pady=20)

    personal_number_label = tk.Label(second_frame, text='PN Pekerja  :')
    personal_number_label.grid(column=0, row=x+1+1, sticky='w', pady=20)
    # personal_number_label.pack(side = LEFT,fill=X, expand=True)

    personal_number = tk.IntVar()
    personal_number_entry = tk.Entry(second_frame, textvariable=personal_number)
    personal_number_entry.grid(column=1, row=x+1+1, sticky='nesw', pady=20)
    # personal_number_entry.pack(side = RIGHT,fill=X, expand=True)

    worker_role_label = tk.Label(second_frame, text='Jabatan  :')
    worker_role_label.grid(column=0, row=x+1+2, sticky='w', pady=20)
    # worker_role_label.pack(side = LEFT,fill=X, expand=True)

    worker_role = tk.StringVar()
    worker_role_entry = tk.Entry(second_frame, textvariable=worker_role)
    worker_role_entry.grid(column=1, row=x+1+2, sticky='nesw', pady=20)
    # worker_role_entry.pack(side = RIGHT,fill=X, expand=True)

    worker_branch_code_label = tk.Label(second_frame, text='Kode Uker  :')
    worker_branch_code_label.grid(column=0, row=x+1+3, sticky='w', pady=20)
    # worker_branch_code_label.pack(side = LEFT,fill=X, expand=True)

    worker_branch_code = tk.StringVar()
    worker_branch_code_entry = tk.Entry(
        second_frame, textvariable=worker_branch_code)
    worker_branch_code_entry.grid(column=1, row=x+1+3, sticky='nesw', pady=20)
    # worker_branch_code_entry.pack(side = RIGHT,fill=X, expand=True)

    # open image

    rows = x+1+4
    condition_label = tk.Label(second_frame, text='Keterangan Kondisi Fisik  :')
    condition_label.grid(column=0, row=rows, sticky='w', pady=20)

    cond_options = ["Baik", "Bermasalah"]
    clicked_cond = tk.StringVar()
    clicked_cond.set("Select")
    tk.OptionMenu(second_frame, clicked_cond, *cond_options).grid(column=1, row=rows)
    rows += 1

    sticker_label = tk.Label(
        second_frame, text='Terdapat Sticker Pengadaan dan BRIBOX  :')
    sticker_label.grid(column=0, row=rows, sticky='w', pady=20)

    sticker_options = ["Ya", "Tidak"]
    clicked_bribox = tk.StringVar()
    clicked_bribox.set("Select")
    tk.OptionMenu(second_frame, clicked_bribox, *
                  sticker_options).grid(column=1, row=rows)
    rows += 1

    for item in image_dict_list:
        attachment_label = tk.Label(second_frame, text=f'Foto {item}  : ')
        attachment_label.grid(column=0, row=rows, sticky='w', pady=20)
        # attachment_label.pack(fill=X, expand=True)

        attachment_button = tk.Button(
            second_frame,
            text=f'Pilih foto {item}',
            command=saveImagePath,
        ).grid(column=1, row=rows, pady=20)
        rows += 1

    additional_info_label = tk.Label(second_frame, text='Informasi Tambahan  :')
    additional_info_label.grid(column=0, row=rows, sticky='w', pady=20)
    # worker_branch_code_label.pack(side = LEFT,fill=X, expand=True)

    text_area = scrolledtext.ScrolledText(
        second_frame, wrap=tk.WORD, width=25, height=5)
    text_area.insert('insert', 'ex: password ditempel di meja kerja, dll.\nIsi \'-\' jika tidak ada', 0)
    text_area.bind("<Button-1>", lambda x: text_area.delete(0.0, tk.END))
    text_area.grid(column=1, row=rows, sticky='w')
    rows += 1

    ttk.Button(
        second_frame,
        text="Submit",
        command=getStartProcedures,
    ).grid(column=1, row=rows, pady=20)

    root.mainloop()
