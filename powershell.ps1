cmd.exe /c 'echo =============================================================================================================================== >> report.txt'
(Get-NetIPAddress -AddressFamily IPV4 -InterfaceAlias "Ethernet*").IPAddress >> report.txt
cmd.exe /c 'echo "CEK APLIKASI" >> report.txt'
cmd.exe /c 'echo =============================================================================================================================== >> report.txt'

cmd.exe /c 'dir "C:\Program Files (x86)" >> report.txt'
cmd.exe /c 'dir "C:\Program Files" >> report.txt'
cmd.exe /c 'echo                                                                                                                                 .>> report.txt'
cmd.exe /c 'echo                                                                                                                                 .>> report.txt'
cmd.exe /c 'echo =============================================================================================================================== >> report.txt'
cmd.exe /c 'echo "CEK OPEN PORT" >> report.txt'
cmd.exe /c 'echo =============================================================================================================================== >> report.txt'

cmd.exe /c "netstat -a | findstr /r /c:"445" /c:"3389" /c:"21" /c:"22" /c:"23" /c:"5900" >> report.txt"
cmd.exe /c 'echo                                                                                                                                 .>> report.txt'
cmd.exe /c 'echo                                                                                                                                 .>> report.txt'
cmd.exe /c 'echo "PROXY" >> report.txt'
cmd.exe /c 'echo =============================================================================================================================== >> report.txt'
Get-ItemProperty -Path "Registry::HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings" >> report.txt

cmd.exe /c 'echo "CEK EMAIL/WEBMAIL" >> report.txt'
cmd.exe /c 'echo =============================================================================================================================== >> report.txt'

cmd.exe /c "ping gmail.com >> report.txt"
cmd.exe /c "ping webmail.bri.co.id >> report.txt"
cmd.exe /c "ping webmail1.bri.co.id >> report.txt"
cmd.exe /c 'echo                                                                                                                                 .>> report.txt'
cmd.exe /c 'echo                                                                                                                                 .>> report.txt'
cmd.exe /c 'echo "CEK FIREWALL SETTING" >> report.txt'
cmd.exe /c 'echo =============================================================================================================================== >> report.txt'


Get-NetFirewallProfile | Select Name, Enabled >> report.txt

cmd.exe /c 'echo "CEK USB" >> report.txt'
cmd.exe /c 'echo =============================================================================================================================== >> report.txt'
cmd.exe /c 'echo "USB Power/Device nyala?.... 3=Enable | 4= Disable" >> report.txt'
 Get-ItemProperty  "HKLM:\SYSTEM\CurrentControlSet\services\USBSTOR" -name start >> report.txt
cmd.exe /c 'echo =============================================================================================================================== >> report.txt'
cmd.exe /c 'echo "USB All storage access.... 3=Enable usb access| 4= Disable usb access" >> report.txt'
 Get-ItemProperty  "HKLM:\Software\Policies\Microsoft\Windows\RemovableStorageDevices\" -name "Deny*" >> report.txt
 cmd.exe /c 'echo =============================================================================================================================== >> report.txt'
 cmd.exe /c 'echo "USB Read/Write.... 3=Enable usb read/write| 4= Disable usb read/write" >> report.txt'
 Get-ItemProperty  "HKLM:\Software\Policies\Microsoft\Windows\RemovableStorageDevices\{53f5630d-b6bf-11d0-94f2-00a0c91efb8b}" -name "Deny*" >> report.txt

cmd.exe /c 'echo "CEK USER" >> report.txt'
cmd.exe /c 'echo =============================================================================================================================== >> report.txt'

Get-LocalUser >> report.txt
Get-CimInstance -ClassName Win32_OperatingSystem >> report.txt

cmd.exe /c 'echo "PASSWORD POLICY" >> report.txt'
cmd.exe /c 'echo =============================================================================================================================== >> report.txt'

cmd.exe /c 'net accounts >> report.txt'
cmd.exe /c 'echo                                                                                                                                 .>> report.txt'
cmd.exe /c 'echo                                                                                                                                 .>> report.txt'
cmd.exe /c 'echo "CEK SCREENSAVER" >> report.txt'
cmd.exe /c 'echo =============================================================================================================================== >> report.txt'
Get-WmiObject -Class Win32_Desktop >> report.txt

cmd.exe /c 'echo                                                                                                                                 .>> report.txt'
cmd.exe /c 'echo '
Get-ComputerInfo -Property "*version" >> report.txt


cmd.exe /c 'echo "CEK ANTIVIRUS" >> report.txt'
cmd.exe /c 'echo =============================================================================================================================== >> report.txt'

cmd.exe /c 'dir "C:\Program Files (x86)" | findstr /r /c:"Trend" /c:"32" /c:"eset" /c:"irus" /c:"afee" /c:"defend" >> report.txt'
cmd.exe /c 'dir "C:\Program Files" | findstr /r /c:"Trend" /c:"32" /c:"eset" /c:"irus" /c:"afee" /c:"defend" >> report.txt'
cmd.exe /c 'echo                                                                                                                                 .>> report.txt'
cmd.exe /c 'echo                                                                                                                                 .>> report.txt'
cmd.exe /c 'echo "CEK EVENT VIEWER SYSTEM" >> report.txt'
cmd.exe /c 'echo =============================================================================================================================== >> report.txt'

Get-EventLog System -Entrytype Error -Before 16/06/2022 -Newest 1 >> report.txt
cmd.exe /c 'echo                                                                                                                                 .>> report.txt'
cmd.exe /c 'echo                                                                                                                                 .>> report.txt'
cmd.exe /c 'echo "CEK EVENT VIEWER SECURITY" >> report.txt'
cmd.exe /c 'echo =============================================================================================================================== >> report.txt'
Get-EventLog Security -Entrytype Error -Before 16/06/2022 -Newest 1 >> report.txt
cmd.exe /c 'echo                                                                                                                                 .>> report.txt'
cmd.exe /c 'echo                                                                                                                                 .>> report.txt'
cmd.exe /c 'echo "RESUME EVENT VIEWER" >> report.txt'
cmd.exe /c 'echo =============================================================================================================================== >> report.txt'
Get-EventLog -List >> report.txt
