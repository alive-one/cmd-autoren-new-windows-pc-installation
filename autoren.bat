rem | Generate pseudorandom number
rem | And took two first digidts for TIMEOUT /T
rem | To ensure that systems will not try to write inv-numbers.txt simultaneously
rem | (Even if one digit will be generated it is not cause error)
rem | After all we have 100 seconds timeout plus random from 0 to 99
TIMEOUT /T 1%RANDOM:~0,2%

rem | Get MAC-address and assign to variable macaddr
@FOR /F %%a IN ('getmac /fo table /nh') DO SET macaddr=%%a

rem | Mount network share as disk Z:
net use Z: \\ushare\backup\autoren

rem | Go to network share
Z: 

rem | Check if there is file with Windows 10 and MS Office 13 keys
rem | If file exists execute activation code
rem | Else skip activation 
IF NOT EXIST Z:\w10-mso13-key.txt GOTO L1

rem | Read Windows 10 and MS Office 13 keys from Z:\w10-mso13-key.txt
rem | Specify letter "r" as eol simbol so that cycle pass lines with comments in file Z:\w10-mso13-key.txt
FOR /F "tokens=1,2 eol=r" %%a IN (Z:\w10-mso13-key.txt) DO (
SET w10key=%%a
SET mso13key=%%b
)

rem | If Windows 10 key is not equal 29 simbols skip 
rem | And move to MS Office 13 activation code
IF %w10key% NEQ 29 GOTO L2

rem | Switch to disk C: 
rem | Otherwise script try to call activation code from Z:
C:

rem | Windows 10 Activation
slmgr.vbs //b /upk 
slmgr.vbs //b /ipk %w10key%
slmgr.vbs //b /ato

rem | Jump here if Windows 10 key is not equal 29 simbols
:L2 

rem | Switch to disk C: 
rem | Otherwise script try to call activation code from Z:
C:

rem | Activate both x64 and x32 versions of MS Office 13
rem | Otherwise activation fail
cd C:\Program Files (x86)\Microsoft Office\Office15
cscript ospp.vbs /inpkey:%mso13key%
cscript ospp.vbs /act

cd C:\Program Files\Microsoft Office\Office15 
cscript ospp.vbs /inpkey:%mso13key%
cscript ospp.vbs /act

rem | Jump here if Z:\w10-mso13-key.txt deos not exist
:L1 

rem | Switch to Z: and start code which rename PC
Z:

rem | Read file Z:\inv-numbers.txt
rem | Use commas and spaces as delimiters 
rem | Assign first token of string to i variable
rem | Rest of file assign to j variable 
rem | Rename PC using i variable as new name
rem | Getting current name from %COMPUTERNAME%

rem | Take mac-адрес from macaddr variable and new PC name (inventary number) from i variable
rem | And add string with this info in file new-pc-names.txt 
rem | also adding capital I before inventary number.

rem | Then rewrite inv-numbers.txt with content of j variable 
rem | As result with every cycle iteration inv-numbers.txt each token, one by one, taken from file and assigned to PCs
rem | And also stored to file new-pc-names.txt alongside with mac address of current PCs

rem | In the end assign i value to pc_name variable so that hwinfo script could store it in html-file
rem | Because current random name of PC will change only after reboot

@FOR /F "tokens=1* delims=, " %%i IN (Z:\inv-numbers.txt) DO (
wmic computersystem where caption='%COMPUTERNAME%' rename I%%i
@ECHO I%%i %macaddr% >> Z:\new-pc-names.txt
@ECHO %%j > Z:\inv-numbers.txt
SET pc_name=%%i
)

rem | Check if html-file with hwinfo already exist, and if not skip writing CSS and jump to config writing directly
rem | TODO: Write hwinfo file in standalone folder
IF EXIST Z:\%pc_name%.html GOTO L3

rem | Write CSS table
@ECHO ^<head^>^<style^> >> Z:\%pc_name%.html
@ECHO .div-main {width: 100%%; height: 800px; padding: 10px; margin: 10px;} >> Z:\%pc_name%.html
@ECHO .div-table {display: inline-block; width: auto; height: 790px; border-left: 1px solid black; padding: 10px; float: left; margin-top: 10px; margin-left: 10px; margin-bottom: 10px;} >> Z:\%pc_name%.html
@ECHO .div-table-sec {display: inline-block; width: auto; height: 790px; border-left: 1px solid black; padding: 10px; margin-top: 10px; margin-bottom: 10px;} >> Z:\%pc_name%.html
@ECHO .div-table-row {display: table-row;} >> Z:\%pc_name%.html
@ECHO .div-table-cell {width: auto; padding: 10px 50px 10px; border-top: 1px solid black; border-left: 1px solid black; float: left;} >> Z:\%pc_name%.html
@ECHO .div-table-cell-zero {width: auto; padding: 10px 50px 10px; border-top: 1px solid black; border-left: 1px solid black; float: left; color: red;} >> Z:\%pc_name%.html
@ECHO .div-table-cell-sec {width: 256px; padding: 10px 50px 10px; border-top: 1px solid black; border-left: 1px solid black; float: left;} >> Z:\%pc_name%.html
@ECHO .div-table-cell-third {width: 128px; padding: 10px 50px 10px; border-top: 1px solid black; border-left: 1px solid black; float: left;} >> Z:\%pc_name%.html
@ECHO ^</head^>^</style^> >> Z:\%pc_name%.html

rem | Jump here if hwinfo file already exists
:L3

rem | Open main positioning div
@ECHO ^<div class=^"div-main^"^> >> Z:\%pc_name%.html

rem | Open first div-table
@ECHO ^<div class=^"div-table^"^> >> Z:\%pc_name%.html

rem | Write local date and time
rem | TODO: Add server date from net time or ntp-server
rem | TODO: Because local date may be incorrect sometimes
@ECHO ^<div class=^"div-table-row^"^>Local Date^</div^> >> Z:\%pc_name%.html
@ECHO ^<div class=^"div-table-row^"^>^<div class=^"div-table-cell-zero^"^>%DATE% %TIME:~0,-3%^</div^>^</div^> >> Z:\%pc_name%.html

rem | Write system name (hostname, computer name, you name it).
rem | Add capital I before pc_name variable value when writing to file 
@ECHO ^<div class=^"div-table-row^"^>Computer Name^</div^> >> Z:\%pc_name%.html
@ECHO ^<div class=^"div-table-row^"^>^<div class=^"div-table-cell^"^>I%pc_name%^</div^>^</div^> >> Z:\%pc_name%.html

rem | Write info on motherboard
rem | Write header for motherboard info
@ECHO ^<div class=^"div-table-row^"^>Motherboard^</div^> >> Z:\%pc_name%.html

rem | Skip two first strings because in CSV format first two strings are "headers" which in fact means empty.
rem | Use separated FOR cycles with CSV output to get baseboard manufacturer and model
rem | Because CSV use comma as delimiter and some motherboards have comma in model name

rem | Write baseboard manufacturer
@FOR /F "skip=2 delims=, tokens=2" %%i IN ('wmic baseboard get Manufacturer /format:csv') DO (
@ECHO ^<div class=^"div-table-row^"^>^<div class=^"div-table-cell^"^>%%i^</div^> >> Z:\%pc_name%.html
) 

rem | Write baseboard model
@FOR /F "skip=2 delims=, tokens=2" %%i IN ('wmic baseboard get Product /format:csv') DO (
@ECHO ^<div class=^"div-table-cell^"^>%%i^</div^>^</div^> >> Z:\%pc_name%.html
)

rem | Write info on CPUs
rem | Write header
@ECHO ^<div class=^"div-table-row^"^>CPU^</div^> >> Z:\%pc_name%.html

rem | Get CPUs info in CSV format using cycle since there may be mode than one CPU
rem | Skip first two strings with headers
rem | Write all string except first token cause in CSV format first token is always "Node"
FOR /F "skip=2 delims=, tokens=1,*" %%i IN ('wmic cpu get name /format:csv') DO (
ECHO ^<div class=^"div-table-row^"^>^<div class=^"div-table-cell^"^>%%j^</div^>^</div^> >> Z:\%pc_name%.html
)

rem | Write info on RAM
rem | Write header
@ECHO ^<div class=^"div-table-row^"^>RAM^</div^> >> Z:\%pc_name%.html

rem | Allow to change variables in cycle for correct represenatation of RAM capacity
Setlocal EnableDelayedExpansion

rem | Write info on each module capacity, slot and speed
@FOR /F "tokens=1,2,3" %%a IN ('wmic memorychip get capacity^,devicelocator^,speed ^| findstr [0-9]') DO (
@ECHO ^<div class=^"div-table-row^"^>^<div class=^"div-table-cell^"^>%%b^</div^> >> Z:\%pc_name%.html

rem | Use inner cycle to calculate RAM capacity in Mb 
rem | Recieve every slot capacity in bytes and divide by 1048576 to get Mb
rem | Use powershell cause CMD can not handle such big numbers
@FOR /F %%a IN ('powershell %%a/1048576') DO (
SET /A mem_fnl=%%a
@ECHO ^<div class=^"div-table-cell^"^>!mem_fnl! Mb^</div^> >> Z:\%pc_name%.html
@ECHO ^<div class=^"div-table-cell^"^>%%c Mhz^</div^>^</div^> >> Z:\%pc_name%.html
)
)

rem | Get info on Storage devices
rem | Write header
@ECHO ^<div class=^"div-table-row^"^>Storage Devices^</div^> >> Z:\%pc_name%.html

rem | Use CSV output cause this way needed values divided by commas
rem | And even space divided parts of string considered as one token what we need in this case
rem | Skip two first header lines
rem | For all storage devices which marked as "fixed hard disk media" get model, size and S.M.A.R.T. status
rem | Use powershell to divide storage value by 1000000000 to get Mb from bytes
@FOR /F "skip=2 delims=, tokens=2-4" %%i IN ('wmic diskdrive where ^(MediaType^="Fixed hard disk media"^) get model^,size^,status /format:csv') DO (
@ECHO ^<div class=^"div-table-row^"^>^<div class=^"div-table-cell^"^>%%i^</div^> >> Z:\%pc_name%.html
@FOR /F %%j IN ('powershell %%j/1000000000') DO (
SET /A stor_fnl=%%j
@ECHO ^<div class=^"div-table-cell^"^>!stor_fnl! Gb^</div^> >> Z:\%pc_name%.html
@ECHO ^<div class=^"div-table-cell^"^>%%k^</div^>^</div^> >> Z:\%pc_name%.html
))

rem | Write header for info on Video Adapters
@ECHO ^<div class=^"div-table-row^"^>Video Adapters^</div^> >> %~dp0%computername%.html

rem | Get currently installed in system video apdapters device IDs
FOR /F "skip=1 tokens=2 delims=&" %%a in ('wmic path win32_VideoController get PNPDeviceID') DO (

rem | Find regedit key where stored current video adapter device ID and assign this regedit key to varaiable %%b
FOR /F %%b IN ('REG QUERY HKLM\SYSTEM\CurrentControlSet\Control\Class\{4D36E968-E325-11CE-BFC1-08002BE10318} /f %%a /s /t REG_SZ ^| find "{"') DO (

rem | UseТеперь в ветке реестра в которой нашли DeviceID ищем название видеокарты
rem | Отбрасываем первые два токена чтобы вывести только название видекарты - оставшаяся часть строки
FOR /F "tokens=1,2*" %%c IN ('REG QUERY ^"%%b^" /f HardwareInformation.AdapterString /t REG_SZ ^| find "REG_SZ"') DO (

rem | Присваиваем переменной строку с именем видеокарты
SET v_name=%%e
)

rem | Ищем во все той же ветке актуальное количество памяти
rem | Используем powershell для пребразования hex значения в дестичное
rem | Делим тоже с помощью powershell чтобы преобразовать в Мб
FOR /F "tokens=3" %%d IN ('REG QUERY ^"%%b^" /f HardwareInformation.qwMemorySize /t REG_QWORD ^| find "0x"') DO (
SET v_mem=powershell [uint64]^('%%d'^)/1048576
)
)
rem | Пишем в таблицу название видеокарты
ECHO ^<div class=^"div-table-row^"^>^<div class=^"div-table-cell^"^>!v_name!^</div^>^ >> %~dp0%computername%.html

rem | Пишем в таблицу количество памяти
ECHO ^<div class=^"div-table-cell^"^> >> %~dp0%computername%.html
!v_mem! >> %~dp0%computername%.html
ECHO Mb^</div^>^</div^> >> %~dp0%computername%.html
)

rem | Disable тключаем возможность изменения переменных внутри цикла.
Setlocal DisableDelayedExpansion

rem | Write header for network cards info
@ECHO ^<div class=^"div-table-row^"^>Network Adapters^</div^> >> Z:\%pc_name%.html

rem | Use cycle for there may be more than one network card
rem | Get string with MAC address and Network Card model
rem | Use undefined q variable to write all tokens except first which is network card name
rem | Then use p variable to write first token, which is MAC address 
@FOR /F "tokens=1*" %%p IN ('wmic NIC where PhysicalAdapter^=true get macaddress^,name ^| findstr [0-9]') DO (
@ECHO ^<div class=^"div-table-row^"^>^<div class=^"div-table-cell-sec^"^>%%q^</div^> >> Z:\%pc_name%.html
@ECHO ^<div class=^"div-table-cell^"^>%%p^</div^>^</div^> >> Z:\%pc_name%.html
)

rem | Close first table with main hardware info
@ECHO ^</div^> >> Z:\%pc_name%.html

rem | Open second table
@ECHO ^<div class=^"div-table-sec^"^> >> Z:\%pc_name%.html

rem | Write info on Operating System
@ECHO ^<div class=^"div-table-row^"^>Operating System^</div^> >> Z:\%pc_name%.html
@ECHO ^<div class=^"div-table-row^"^>^<div class=^"div-table-cell^"^> >> Z:\%pc_name%.html
wmic os get caption | findstr "Windows" >> Z:\%pc_name%.html
@ECHO ^</div^> >> Z:\%pc_name%.html
@ECHO ^<div class=^"div-table-cell^"^> >> Z:\%pc_name%.html
wmic os get version | findstr [0-9] >> Z:\%pc_name%.html
@ECHO ^</div^>^</div^> >> Z:\%pc_name%.html

rem | Get info on Network Connections
rem | For each network card get MAC Address and Connection Name
@FOR /F "skip=2 delims=, tokens=2,3" %%i IN ('wmic nic where PhysicalAdapter^=true get MACAddress^, NetConnectionID /format:csv') DO (

rem | For each uniq MAC Address get Connection Name and write it as header
@ECHO  ^<div class=^"div-table-row^"^>%%j^</div^> >> Z:\%pc_name%.html

rem | For each Connection Name write MAC address
@ECHO ^<div class=^"div-table-row^"^>^<div class=^"div-table-cell-third^"^>MAC Address^</div^>^<div class=^"div-table-cell^"^>%%i^</div^>^</div^> >> Z:\%pc_name%.html

rem | For each "Connection Name and MAC address" pair get IP Address, Gateway and Subnet Mask
@FOR /F "skip=2 delims=,{} tokens=2,3,4" %%a IN ('wmic nicconfig where ^(ipenabled^="true" AND macaddress^="%%i"^) get DefaultIPGateway^, IPAddress^, IPSubnet /format:csv') DO (

rem | Get rid of IPv6 address if there any
rem | Use inner cycle to sort variable b with semicolon as delimiters
rem | because IPv4 and IPv6 addresses in this case separated by semicolon
rem | And write first token (IPv4 address) to file
@FOR /F "delims=; tokens=1" %%z IN (^"%%b^") DO (
@ECHO ^<div class=^"div-table-row^"^>^<div class=^"div-table-cell-third^"^>IP Address^</div^>^<div class=^"div-table-cell^"^>%%z^</div^>^</div^> >> Z:\%pc_name%.html
)

rem | Get rid of 64 bitmask same way
@FOR /F "delims=; tokens=1" %%y IN (^"%%c^") DO (
@ECHO ^<div class=^"div-table-row^"^>^<div class=^"div-table-cell-third^"^>Subnet^</div^> ^<div class=^"div-table-cell^"^>%%y^</div^>^</div^> >> Z:\%pc_name%.html
)
@ECHO ^<div class=^"div-table-row^"^>^<div class=^"div-table-cell-third^"^>Gateway^</div^> ^<div class=^"div-table-cell^"^>%%a^</div^>^</div^> >> Z:\%pc_name%.html
)
)

rem | Close second positioning div
@ECHO ^</div^> >> Z:\%pc_name%.html

rem | Close main DIV
@ECHO ^</div^> >> Z:\%pc_name%.html
 
rem | Unmount network drive
net use /del Z: /y

rem | Switch to disk C
rem | Otherwise Windows will try to launch command files from disk Z:
C:

rem | Mark folder "Scripts" for deletion as unnecessary
Reg Add HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\RunOnce /V RemoveFiles /D "cmd /C RD /S /Q C:\Windows\Setup\Scripts" /F

rem | Shut down PC (Also we can reboot from here)
shutdown /s /t 0
