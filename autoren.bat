rem | TODO: ����� ���� ������ ��� ������ � html-����, ��� ����������� � ����������� �� ����� ������
rem | ����, ���������� � html-����� ��� ���������� ���� � �� ����� ������ � �� �� mac-������

rem | ���������� ��������������� ����� �� 1 �� 32767 
rem | � ����� ������ ��� ����� ����� ���������� � TIMEOUT /T ��� ������������ ������ �������.
rem | ����� ��� ������ �� ���������� � ����� � �������������� ������������.
rem | (���� ���� ������������� ���� �����, ��� �� ������� ������).
rem | ���� /NOBREAK �� ���� �������� ������� �������� ����� �������.
rem | �� ������ ����� ������� � 100 ������ ���� ��������� �� 0 �� 99
TIMEOUT /T 1%RANDOM:~0,2%

rem ������� mac-����� ���������� � ����������� ���������� macaddr
@FOR /F %%a IN ('getmac /fo table /nh') DO SET macaddr=%%a

rem ��������� ������� ����
net use Z: \\ushare\backup\autoren

rem ������� �� ���� Z:
Z: 

rem | ��������� ���� �� ���� � ������� ����� � �����
rem | ���� ���� ��������� ��� ��������� ����� � �����
rem | ���� ��� ������� ��������� � ���� ������ �� L2
rem | ���� �� ���� ���������� ����� ���� � ������� ������ �� ������ ������ ������
rem | �� ���� ����� � ���� ���������� ������������� � ������� �������
rem | �� ���������� ���� ���������. �������� ���������.
IF NOT EXIST Z:\w10-mso13-key.txt GOTO L1

rem | ������ ����� ����� � ����� �� ����� � ������� � ����������������� ����� Z
rem | ��������� � �������� ������� ��� �������� ����� ���� r 
rem | ����� ���� �� ����� ����������� �� ����� � ������� � ����� ������ �����
FOR /F "tokens=1,2 eol=r" %%a IN (Z:\w10-mso13-key.txt) DO (
SET w10key=%%a
SET mso13key=%%b
)

rem | ����� ������� ��������� ERRORLEVEL ������� ������������ ��� ��������� ����� � ���� �� ��������� �� ������ �� �������
rem | ���-�� ����� rem if errorlevel 1 echo Key is wrong
rem | ��������� � ���������
rem | https://coderoad.ru/34987885/%D0%9A%D0%B0%D0%BA%D0%BE%D0%B2%D1%8B-%D0%B7%D0%BD%D0%B0%D1%87%D0%B5%D0%BD%D0%B8%D1%8F-ERRORLEVEL-%D0%B7%D0%B0%D0%B4%D0%B0%D0%BD%D0%BD%D1%8Brem %D0%B5-%D0%B2%D0%BD%D1%83%D1%82%D1%80%D0%B5%D0%BD%D0%BD%D0%B8%D0%BC%D0%B8-%D0%BA%D0%BE%D0%BC%D0%B0%D0%BD%D0%B4%D0%B0%D0%BC%D0%B8-cmd-exe
rem | �� ��� �� ����� ���� �� ����� errorlevel ����������� ��� ������� � �������� ��� ������. 
rem | ������� ����� ������������� ������ ��� �������� � �� �������������

IF %w10key% NEQ 29 GOTO L2

rem | ��������� �� ���� C: ����� ����� ���������� ������� ��������� � ����� Z
C:

rem | ���������� �����
slmgr.vbs //b /upk 
slmgr.vbs //b /ipk %w10key%
slmgr.vbs //b /ato

:L2 rem | ���� ������� ���� ���� ����� �� ����� 29 ��������
rem | ��������� �� ���� C: ����� ����� ���������� ������� ��������� � ����� Z
C:

rem | ���������� � 64 � 32 ��������� ������ 
rem | ��� ���� ���� ���� ����� ������ �64 �����
cd C:\Program Files (x86)\Microsoft Office\Office15
cscript ospp.vbs /inpkey:%mso13key%
cscript ospp.vbs /act

cd C:\Program Files\Microsoft Office\Office15 
cscript ospp.vbs /inpkey:%mso13key%
cscript ospp.vbs /act

:L1 rem | ���� ������� ���� ��� ����� Z:\w10-mso13-key.txt � ������� ����� � �����

rem | ������� �� ������� ���� � �������� ������ �� �������������� ������
Z:

rem | ������ ���� Z:\inv-numbers.txt
rem | ��������� ������ c �������������� �� ���������, ������ ������������� ������� � �������, � �������� � ����.
rem | �������� ������ ��������� ����������� ���������� i
rem | �������� ���� ��������� �������� ������������� ������� ���������� j
rem | ��������������� ��������� ������ �������� wmic ��������� ��� ����� ������ �������� ���������� i
rem | � ������� ��� ��� ��������� caption �������� �� ���������� ���������� %COMPUTERNAME%

rem | ����� mac-����� ����� �� ���������� macaddr, � ����������� �� ���������� i
rem | � ���������� � ���� new-pc-names.txt ������ "mac-����� - ����������� �����"
rem | �������� I ����� ����������� �������

rem | ����� �������������� ���� inv-numbers.txt ���������� ���������� j
rem | � ���������� ���� inv-numbers.txt ����������� �� ���� ����������� ����� 
rem | � ���� �� �������� "mac-����� - ����������� �����" ������������ ������� "mac- ����� - ����������� �����" ��������� ������
rem | ��� ����������� �������� i ���������� pc_name ����� ����� ������ �������� �� ����� ���� � ������
rem | ���� ������� ��������� ��� ������ �������� �� ���������� ������ ����� ������

@FOR /F "tokens=1* delims=, " %%i IN (Z:\inv-numbers.txt) DO (
wmic computersystem where caption='%COMPUTERNAME%' rename I%%i
@ECHO I%%i %macaddr% >> Z:\new-pc-names.txt
@ECHO %%j > Z:\inv-numbers.txt
SET pc_name=%%i
)

rem | ��������� ���� �� ��� ����� ���� � ���� ���� �� ������� ������ �� �����, � ����� ������� � ������ ������������
rem | TODO ������� ����� ��������� ����� �� ���� ������� hwinfo � ���� ���������� ����.
IF EXIST Z:\%pc_name%.html GOTO L3

rem | ����� ������� ������
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

:L3 rem | ���� ������� ���� ���� � ������ ������ ��� ����������, ����� �� ������ ������ ������� ������. 

rem | ����� div ��� ����������������
@ECHO ^<div class=^"div-main^"^> >> Z:\%pc_name%.html

rem | ��������� div-�������
@ECHO ^<div class=^"div-table^"^> >> Z:\%pc_name%.html

rem | ����� ���� � �����. 
rem | ����� ������� ��� ��������� ���� ��������. ����� �������� �� �����.
rem TODO: �������� ���� � ������� ������� ����� �������� ����� net time. 
rem TODO: ��������� ��������� ���� ����� ���� ������ ����� ����� ��� ������ ��������� ��� � ���� ���.
@ECHO ^<div class=^"div-table-row^"^>Local Date^</div^> >> Z:\%pc_name%.html
@ECHO ^<div class=^"div-table-row^"^>^<div class=^"div-table-cell-zero^"^>%DATE% %TIME:~0,-3%^</div^>^</div^> >> Z:\%pc_name%.html

rem | ����� ��� �������
rem | ��������� � pc_name ������ ����� �� ����� ������� ��� I ����� pc_name
@ECHO ^<div class=^"div-table-row^"^>Computer Name^</div^> >> Z:\%pc_name%.html
@ECHO ^<div class=^"div-table-row^"^>^<div class=^"div-table-cell^"^>I%pc_name%^</div^>^</div^> >> Z:\%pc_name%.html

rem | ����� ���� � �������c��� �����
rem | ����� ��������� ��� ���������� � ����������� �����
@ECHO ^<div class=^"div-table-row^"^>Motherboard^</div^> >> Z:\%pc_name%.html

rem | ���������� ��� ������ ������ (��� ������ ������ � ���������) � � ������ CSV 
rem | ������� ���������� ������� ��� ����������� �������� �������� � ���� �� �������� �������� ������ ��������� (��� �����)
rem | �� ���������� ��� ��������� ������ � ������������� ���� ���� 
rem | ��������� � ������� ��������� ��������� ���� �������, ��� �������� � ������������� ���������� �������

rem | ����� �������������
@FOR /F "skip=2 delims=, tokens=2" %%i IN ('wmic baseboard get Manufacturer /format:csv') DO (
@ECHO ^<div class=^"div-table-row^"^>^<div class=^"div-table-cell^"^>%%i^</div^> >> Z:\%pc_name%.html
) 

rem | ����� ������ ����������� ����� � �� �� ������� ���� ������ � �������������
@FOR /F "skip=2 delims=, tokens=2" %%i IN ('wmic baseboard get Product /format:csv') DO (
@ECHO ^<div class=^"div-table-cell^"^>%%i^</div^>^</div^> >> Z:\%pc_name%.html
)

rem | ����� ���� � ����������.
rem | ����� ��������� ��� ������� � ���� � ����������
@ECHO ^<div class=^"div-table-row^"^>CPU^</div^> >> Z:\%pc_name%.html

rem | �������� ���� � ���������� ������� ������������ � ������� CSV ����������� ��������
rem | ����� ��������� ������� ��� ����������� ������� ��� ���������� ������ ������ ����� ������� ������
rem | �� ���� ��� ��� ���������� � ����������, ��������� ������ ����� � CSV ��� Node �� ���� ��� ������.
FOR /F "skip=2 delims=, tokens=1,*" %%i IN ('wmic cpu get name /format:csv') DO (
ECHO ^<div class=^"div-table-row^"^>^<div class=^"div-table-cell^"^>%%j^</div^>^</div^> >> Z:\%pc_name%.html
)

rem | ����� ���� �� ����������� ������
rem | ����� ��������� �������
@ECHO ^<div class=^"div-table-row^"^>RAM^</div^> >> Z:\%pc_name%.html

rem | ���������� �������� ���������� ������ ����� ��� ���������� ���������� ������ ��� ������ �����
Setlocal EnableDelayedExpansion

rem | ����� ���� � �����, ������� � ��������
@FOR /F "tokens=1,2,3" %%a IN ('wmic memorychip get capacity^,devicelocator^,speed ^| findstr [0-9]') DO (
@ECHO ^<div class=^"div-table-row^"^>^<div class=^"div-table-cell^"^>%%b^</div^> >> Z:\%pc_name%.html

rem | �� ��������� ����� ���������� �������� ���������� � ������� ������� ����� ��� � ������
rem | � ����� �� 1048576 ����� �������� �������� � ����������. 
rem | ���������� powershell ��� ������� ��������� cmd ����� ����� �� ���������.
@FOR /F %%a IN ('powershell %%a/1048576') DO (
SET /A mem_fnl=%%a
@ECHO ^<div class=^"div-table-cell^"^>!mem_fnl! Mb^</div^> >> Z:\%pc_name%.html
@ECHO ^<div class=^"div-table-cell^"^>%%c Mhz^</div^>^</div^> >> Z:\%pc_name%.html
)
)

rem | �������� ���������� � ����������� 
rem | ��������� ������ ������� � ����� ��������� ��� ���������� � ������������
@ECHO ^<div class=^"div-table-row^"^>Storage Devices^</div^> >> Z:\%pc_name%.html

rem | ���������� ����� � ������� csv ��������� ��� ������ ��� �������� ��������� �������� 
rem | � ���� �������� � ��������� ����� ����� ��� ��������� ������, ��� ��� � ����� � ������ � ��������� ������.
rem | ��������� ��� ������ ��������� ����� CSV ����� ��� ������ ������ ������ ���������.
rem | ��� ���� ��������� ��������, ������� ������� �, �������� � ����� ������ � ������ � ����� � �������
rem | ����� ������ � ������� powershell ����� �� 1000000000 ����� �������� � ����� ������������� ��� - � ���������.
@FOR /F "skip=2 delims=, tokens=2-4" %%i IN ('wmic diskdrive where ^(MediaType^="Fixed hard disk media"^) get model^,size^,status /format:csv') DO (
@ECHO ^<div class=^"div-table-row^"^>^<div class=^"div-table-cell^"^>%%i^</div^> >> Z:\%pc_name%.html
@FOR /F %%j IN ('powershell %%j/1000000000') DO (
SET /A stor_fnl=%%j
@ECHO ^<div class=^"div-table-cell^"^>!stor_fnl! Gb^</div^> >> Z:\%pc_name%.html
@ECHO ^<div class=^"div-table-cell^"^>%%k^</div^>^</div^> >> Z:\%pc_name%.html
))

rem | ��������� ����������� �������� ���������� ������ �����.
Setlocal DisableDelayedExpansion

rem | �������� ���������� � �����������, �� ���� ����� ���� ��������� ������� � �����
rem | AdapterRam �� ����������, ��������� ��� ������ ������ ������ 4Gb ��� ��� ����� ������ ������ 4Gb

rem | ������������ AdapterRam ��� ����������� �������� ���-�� ������ ��� 4Gb �����
rem | ������ ���� ���� ��� ����� ����-�� ������ � ����� �����������
rem | ��������� �� 32-� ���������, �� ������, ��������, ��� 6�� �������� 4�� � ����� 2��
rem | �� ����, ����� ������ �� ��������� ����������� � ����� powershell'�� ����������

rem | ��������� ��� ���� ������� ��� ������ ����� ��� ������, � ���� �������� � �����������
rem | ���������� ��������� tokens=* ��� ������ ���� ������� � ������, �� ���� ���� ������ �������. 
@ECHO ^<div class=^"div-table-row^"^>Video Adapters^</div^> >> Z:\%pc_name%.html
@FOR /F "skip=1 tokens=*" %%m IN ('wmic path win32_VideoController get Name ^| findstr "."') DO (
@ECHO ^<div class=^"div-table-row^"^>^<div class=^"div-table-cell^"^>%%m^</div^>^</div^> >> Z:\%pc_name%.html
)

rem | ����� ��������� ��� ���������� � ���������
@ECHO ^<div class=^"div-table-row^"^>Network Adapters^</div^> >> Z:\%pc_name%.html

rem | ��������� ������� ����� ���� ���������, ����� ���������� ����.
rem | �������� ������ � ����� � ������� �������� 
rem | ����� ������� ������� � ������� ������� ���������� %%q ��� ���������� ������ ������, ����� �������, ��� ��� ��� �������� ��������
rem | ����� � ������� ���������� %%p ������� ������ ����� - ��� ��� 
@FOR /F "tokens=1*" %%p IN ('wmic NIC where PhysicalAdapter^=true get macaddress^,name ^| findstr [0-9]') DO (
@ECHO ^<div class=^"div-table-row^"^>^<div class=^"div-table-cell-sec^"^>%%q^</div^> >> Z:\%pc_name%.html
@ECHO ^<div class=^"div-table-cell^"^>%%p^</div^>^</div^> >> Z:\%pc_name%.html
)

rem | ��������� ������� � �������� ����� � ������ (����� ��� ���� ��������� � ����� �����).
@ECHO ^</div^> >> Z:\%pc_name%.html

rem | ��������� ������ ������� ���������������� 
@ECHO ^<div class=^"div-table-sec^"^> >> Z:\%pc_name%.html

rem | ����� ���������� �� ��
@ECHO ^<div class=^"div-table-row^"^>Operating System^</div^> >> Z:\%pc_name%.html
@ECHO ^<div class=^"div-table-row^"^>^<div class=^"div-table-cell^"^> >> Z:\%pc_name%.html
wmic os get caption | findstr "Windows" >> Z:\%pc_name%.html
@ECHO ^</div^> >> Z:\%pc_name%.html
@ECHO ^<div class=^"div-table-cell^"^> >> Z:\%pc_name%.html
wmic os get version | findstr [0-9] >> Z:\%pc_name%.html
@ECHO ^</div^>^</div^> >> Z:\%pc_name%.html

rem | �������� ���� � ������� ������������
rem | � ����� ��� ������ �������� ��� ����������� ���������� �������� ��� � �������� ����������
@FOR /F "skip=2 delims=, tokens=2,3" %%i IN ('wmic nic where PhysicalAdapter^=true get MACAddress^, NetConnectionID /format:csv') DO (

rem | ��������� �� ���������� ����� � ����� �� ������ ��� �������� ���������� ��� ��������� �������
@ECHO  ^<div class=^"div-table-row^"^>%%j^</div^> >> Z:\%pc_name%.html

rem | ��� ������� ���������� ����� ���
@ECHO ^<div class=^"div-table-row^"^>^<div class=^"div-table-cell-third^"^>MAC Address^</div^>^<div class=^"div-table-cell^"^>%%i^</div^>^</div^> >> Z:\%pc_name%.html

rem | ��� ������� ���������� ����� ����, IP �����, ����� �������
@FOR /F "skip=2 delims=,{} tokens=2,3,4" %%a IN ('wmic nicconfig where ^(ipenabled^="true" AND macaddress^="%%i"^) get DefaultIPGateway^, IPAddress^, IPSubnet /format:csv') DO (

rem | ����������� �� IPv6 ������ ���� �� ����������� � ������. 
rem | ���������� � ����� �������� ���������� � ����������� �� IP ������ ��������� ; ��� ��������
rem | ��������� ������� nicconfig ������� ������ ����������� ��������, � ��������� ��������� ������ � �������
rem | �� ��� �������� �������� ��������� ���������� ��� �������� ��� ��� ����� � �������
@FOR /F "delims=; tokens=1" %%z IN (^"%%b^") DO (
@ECHO ^<div class=^"div-table-row^"^>^<div class=^"div-table-cell-third^"^>IP Address^</div^>^<div class=^"div-table-cell^"^>%%z^</div^>^</div^> >> Z:\%pc_name%.html
)

rem | ����������� �� ����������� ����� �������� ��������.
@FOR /F "delims=; tokens=1" %%y IN (^"%%c^") DO (
@ECHO ^<div class=^"div-table-row^"^>^<div class=^"div-table-cell-third^"^>Subnet^</div^> ^<div class=^"div-table-cell^"^>%%y^</div^>^</div^> >> Z:\%pc_name%.html
)
@ECHO ^<div class=^"div-table-row^"^>^<div class=^"div-table-cell-third^"^>Gateway^</div^> ^<div class=^"div-table-cell^"^>%%a^</div^>^</div^> >> Z:\%pc_name%.html
)
)

rem | ��������� ������ ������� ���������������
@ECHO ^</div^> >> Z:\%pc_name%.html

rem | ��������� ������� ������� div ���������������
@ECHO ^</div^> >> Z:\%pc_name%.html
 
rem | ������������ ������� ����
net use /del Z: /y

rem | ��������� �� ���� C:
rem | ����� Windows �������� ��������� ��� ��������� ����� ���� � ����������������� ��� ����� Z:
C:

rem | �������� �� �������� ����� Scripts ��� ��������
Reg Add HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\RunOnce /V RemoveFiles /D "cmd /C RD /S /Q C:\Windows\Setup\Scripts" /F

rem | ��������� ������ 
shutdown /s /t 0