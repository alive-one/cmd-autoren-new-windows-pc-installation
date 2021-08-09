rem | Generate pseudorandom number
rem | And took two first digidts for TIMEOUT /T
rem | To ensure that no systems will use inv-numbers.txt simultaneously
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

:L2 rem | Jump here if Windows 10 key is not equal 29 simbols
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

:L1 rem | Jump here if Z:\w10-mso13-key.txt deos not exist

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

rem | После ПЕРЕЗАПИСЫВАЕМ файл inv-numbers.txt содержимым переменной j
rem | В результате файл inv-numbers.txt уменьшается на один инвентарный номер 
rem | А файл со связками "mac-адрес - инвентарный номер" дописывается связкой "mac- адрес - инвентарный номер" локальной машины
rem | Еще присваиваем значение i переменной pc_name чтобы потом читать скриптом по сбору инфы о железе
rem | Ведь текущее рандомное имя машины сменится на правильное только после ребута

@FOR /F "tokens=1* delims=, " %%i IN (Z:\inv-numbers.txt) DO (
wmic computersystem where caption='%COMPUTERNAME%' rename I%%i
@ECHO I%%i %macaddr% >> Z:\new-pc-names.txt
@ECHO %%j > Z:\inv-numbers.txt
SET pc_name=%%i
)

rem | Проверяем есть ли уже такой файл и если есть то таблицу стилей не пишем, а сразу прыгаем к записи конфигурации
rem | TODO сделать потом отдельную папку на шаре назвать hwinfo и туда складывать инфу.
IF EXIST Z:\%pc_name%.html GOTO L3

rem | Пишем таблицу стилей
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

:L3 rem | Сюда прыгаем если файл с именем машины уже существует, чтобы не писать дважды таблицу стилей. 

rem | Пишем div для позиционирования
@ECHO ^<div class=^"div-main^"^> >> Z:\%pc_name%.html

rem | Открываем div-таблицу
@ECHO ^<div class=^"div-table^"^> >> Z:\%pc_name%.html

rem | Пишем дату и время. 
rem | Время выводим без последних трех символов. Такая точность не нужна.
rem TODO: Добавить дату с сервера которую можно получать через net time. 
rem TODO: Поскольку локальная дата может быть вообще любой ввиду как севшей батарейки так и чего еще.
@ECHO ^<div class=^"div-table-row^"^>Local Date^</div^> >> Z:\%pc_name%.html
@ECHO ^<div class=^"div-table-row^"^>^<div class=^"div-table-cell-zero^"^>%DATE% %TIME:~0,-3%^</div^>^</div^> >> Z:\%pc_name%.html

rem | Пишем имя системы
rem | Поскольку в pc_name только номер то пишем сначала еще I перед pc_name
@ECHO ^<div class=^"div-table-row^"^>Computer Name^</div^> >> Z:\%pc_name%.html
@ECHO ^<div class=^"div-table-row^"^>^<div class=^"div-table-cell^"^>I%pc_name%^</div^>^</div^> >> Z:\%pc_name%.html

rem | Пишем инфо о материнcкой плате
rem | Пишем заголовок для информации о материнской плате
@ECHO ^<div class=^"div-table-row^"^>Motherboard^</div^> >> Z:\%pc_name%.html

rem | Пропускаем две первые строки (это пустая строка и заголовки) и в фомате CSV 
rem | который использует запятые как разделители подстрок передаем в цикл из которого забираем нужную подстроку (или токен)
rem | Не используем для получения модели и производителя один цикл 
rem | Поскольку в моделях некоторых материнок есть запятая, что приводит к неправильному разделению токенов

rem | Пишем производителя
@FOR /F "skip=2 delims=, tokens=2" %%i IN ('wmic baseboard get Manufacturer /format:csv') DO (
@ECHO ^<div class=^"div-table-row^"^>^<div class=^"div-table-cell^"^>%%i^</div^> >> Z:\%pc_name%.html
) 

rem | Пишем модель материнской платы в ту же таблицу куда писали и производителя
@FOR /F "skip=2 delims=, tokens=2" %%i IN ('wmic baseboard get Product /format:csv') DO (
@ECHO ^<div class=^"div-table-cell^"^>%%i^</div^>^</div^> >> Z:\%pc_name%.html
)

rem | Пишем инфо о процессоре.
rem | Пишем заголовок для таблицы с инфо о процессоре
@ECHO ^<div class=^"div-table-row^"^>CPU^</div^> >> Z:\%pc_name%.html

rem | Получаем инфо о процессоре которую представляем в формате CSV разделенном запятыми
rem | Затем используя запятые как разделители выводим все оставшиеся токены строки кроме первого токена
rem | То есть как раз информацию о процессоре, поскольку первый токен в CSV это Node то есть имя машины.
FOR /F "skip=2 delims=, tokens=1,*" %%i IN ('wmic cpu get name /format:csv') DO (
ECHO ^<div class=^"div-table-row^"^>^<div class=^"div-table-cell^"^>%%j^</div^>^</div^> >> Z:\%pc_name%.html
)

rem | Пишем инфо об оперативной памяти
rem | Пишем заголовок таблицы
@ECHO ^<div class=^"div-table-row^"^>RAM^</div^> >> Z:\%pc_name%.html

rem | Разрешщаем изменять переменные внутри цикла для коректного вычисления объема ОЗУ внутри цикла
Setlocal EnableDelayedExpansion

rem | Пишем инфо о слоте, емкости и скорости
@FOR /F "tokens=1,2,3" %%a IN ('wmic memorychip get capacity^,devicelocator^,speed ^| findstr [0-9]') DO (
@ECHO ^<div class=^"div-table-row^"^>^<div class=^"div-table-cell^"^>%%b^</div^> >> Z:\%pc_name%.html

rem | Во вложенном цикле перебираем значения переменной с объемом каждого слота ОЗУ в байтах
rem | и делим на 1048576 чтобы получить значение в мегабайтах. 
rem | Используем powershell для деления поскольку cmd такие числа не осиливает.
@FOR /F %%a IN ('powershell %%a/1048576') DO (
SET /A mem_fnl=%%a
@ECHO ^<div class=^"div-table-cell^"^>!mem_fnl! Mb^</div^> >> Z:\%pc_name%.html
@ECHO ^<div class=^"div-table-cell^"^>%%c Mhz^</div^>^</div^> >> Z:\%pc_name%.html
)
)

rem | Получаем информацию о накопителях 
rem | Открываем строку таблицы и пишем заголовок для подраздела с накопителями
@ECHO ^<div class=^"div-table-row^"^>Storage Devices^</div^> >> Z:\%pc_name%.html

rem | Испольщуем вывод в формате csv поскольку так нужные нам значения разделены запятыми 
rem | И даже названия с пробелами можно брать как одинарные токены, что нам и нужно в случае с названием дисков.
rem | Проускаем две строки поскольку вывод CSV пишет еще пустую строку помимо заголовка.
rem | Для всех устройств хранения, которые система и, получаем в цикле модель и размер и пишем в таблицу
rem | Затем размер с помощью powershell делим на 1000000000 чтобы привести в более удобочитаемый вид - в гигабайты.
@FOR /F "skip=2 delims=, tokens=2-4" %%i IN ('wmic diskdrive where ^(MediaType^="Fixed hard disk media"^) get model^,size^,status /format:csv') DO (
@ECHO ^<div class=^"div-table-row^"^>^<div class=^"div-table-cell^"^>%%i^</div^> >> Z:\%pc_name%.html
@FOR /F %%j IN ('powershell %%j/1000000000') DO (
SET /A stor_fnl=%%j
@ECHO ^<div class=^"div-table-cell^"^>!stor_fnl! Gb^</div^> >> Z:\%pc_name%.html
@ECHO ^<div class=^"div-table-cell^"^>%%k^</div^>^</div^> >> Z:\%pc_name%.html
))

rem | Отключаем возможность изменять переменные внутри цикла.
Setlocal DisableDelayedExpansion

rem | Получаем информацию о видеокартах, их ведь может быть несколько поэтому в цикле
rem | AdapterRam не используем, поскольку при объеме памяти больше 4Gb оно все равно выдаст только 4Gb

rem | Использовать AdapterRam для определения бОльшего кол-ва памяти чем 4Gb можно
rem | Просто надо весь его вывод куда-то писать и потом суммировать
rem | Поскольку он 32-х разрядный, он выдает, например, для 6Гб карточки 4Гб и потом 2Гб
rem | То есть, можно писать во временный текстовичок и потом powershell'ом складывать

rem | Поскольку нам надо вывести для каждой карты всю строку, а цикл работает с подстроками
rem | Используем директиву tokens=* для вывода всех токенов в строке, то есть всей строки целиком. 
@ECHO ^<div class=^"div-table-row^"^>Video Adapters^</div^> >> Z:\%pc_name%.html
@FOR /F "skip=1 tokens=*" %%m IN ('wmic path win32_VideoController get Name ^| findstr "."') DO (
@ECHO ^<div class=^"div-table-row^"^>^<div class=^"div-table-cell^"^>%%m^</div^>^</div^> >> Z:\%pc_name%.html
)

rem | Пишем заголовок для информации о сетевухах
@ECHO ^<div class=^"div-table-row^"^>Network Adapters^</div^> >> Z:\%pc_name%.html

rem | Поскольку сетевух может быть несколько, опять используем цикл.
rem | Получаем строку с маком и моделью сетевухи 
rem | Затем сначала выводим с помощью неявной переменной %%q все оставшиеся токены строки, кроме первого, это как раз название сетевухи
rem | Затем с помощью переменной %%p выводим первый токен - это мак 
@FOR /F "tokens=1*" %%p IN ('wmic NIC where PhysicalAdapter^=true get macaddress^,name ^| findstr [0-9]') DO (
@ECHO ^<div class=^"div-table-row^"^>^<div class=^"div-table-cell-sec^"^>%%q^</div^> >> Z:\%pc_name%.html
@ECHO ^<div class=^"div-table-cell^"^>%%p^</div^>^</div^> >> Z:\%pc_name%.html
)

rem | Закрываем таблицу с основной инфой о железе (потом это надо перенести в самый конец).
@ECHO ^</div^> >> Z:\%pc_name%.html

rem | Открываем вторую таблицу позиционирования 
@ECHO ^<div class=^"div-table-sec^"^> >> Z:\%pc_name%.html

rem | Пишем информацию об ОС
@ECHO ^<div class=^"div-table-row^"^>Operating System^</div^> >> Z:\%pc_name%.html
@ECHO ^<div class=^"div-table-row^"^>^<div class=^"div-table-cell^"^> >> Z:\%pc_name%.html
wmic os get caption | findstr "Windows" >> Z:\%pc_name%.html
@ECHO ^</div^> >> Z:\%pc_name%.html
@ECHO ^<div class=^"div-table-cell^"^> >> Z:\%pc_name%.html
wmic os get version | findstr [0-9] >> Z:\%pc_name%.html
@ECHO ^</div^>^</div^> >> Z:\%pc_name%.html

rem | Получаем инфу о сетевых подключениях
rem | В цикле для каждой сетевухи как физического устройства получаем мак и название соединения
@FOR /F "skip=2 delims=, tokens=2,3" %%i IN ('wmic nic where PhysicalAdapter^=true get MACAddress^, NetConnectionID /format:csv') DO (

rem | Фильтруем по полученным макам и пишем на каждый мак название соединения как заголовок таблицы
@ECHO  ^<div class=^"div-table-row^"^>%%j^</div^> >> Z:\%pc_name%.html

rem | Для каждого соединения пишем мак
@ECHO ^<div class=^"div-table-row^"^>^<div class=^"div-table-cell-third^"^>MAC Address^</div^>^<div class=^"div-table-cell^"^>%%i^</div^>^</div^> >> Z:\%pc_name%.html

rem | ДЛя каждого соединения пишем шлюз, IP адрес, маску подсети
@FOR /F "skip=2 delims=,{} tokens=2,3,4" %%a IN ('wmic nicconfig where ^(ipenabled^="true" AND macaddress^="%%i"^) get DefaultIPGateway^, IPAddress^, IPSubnet /format:csv') DO (

rem | Избавляемся от IPv6 адреса если он представлен в выводе. 
rem | Перебираем в цикле значения переменной с информацией об IP адресе используя ; как делитель
rem | поскольку команда nicconfig выводит токены разделенные запятыми, а подтокены разделены точкой с запятой
rem | то для перебора значений подтокена используем как делитель как раз точку с запятой
@FOR /F "delims=; tokens=1" %%z IN (^"%%b^") DO (
@ECHO ^<div class=^"div-table-row^"^>^<div class=^"div-table-cell-third^"^>IP Address^</div^>^<div class=^"div-table-cell^"^>%%z^</div^>^</div^> >> Z:\%pc_name%.html
)

rem | Избавляемся от разрядности маски подобным способом.
@FOR /F "delims=; tokens=1" %%y IN (^"%%c^") DO (
@ECHO ^<div class=^"div-table-row^"^>^<div class=^"div-table-cell-third^"^>Subnet^</div^> ^<div class=^"div-table-cell^"^>%%y^</div^>^</div^> >> Z:\%pc_name%.html
)
@ECHO ^<div class=^"div-table-row^"^>^<div class=^"div-table-cell-third^"^>Gateway^</div^> ^<div class=^"div-table-cell^"^>%%a^</div^>^</div^> >> Z:\%pc_name%.html
)
)

rem | Закрываем вторую таблицу позицонирования
@ECHO ^</div^> >> Z:\%pc_name%.html

rem | Закрываем внешний главный div позицонирования
@ECHO ^</div^> >> Z:\%pc_name%.html
 
rem | Размонтируем сетевой диск
net use /del Z: /y

rem | Переходим на диск C:
rem | Иначе Windows пытается запустить все командные файлы ниже с размонтированного уже диска Z:
C:

rem | Помечаем на удаление папку Scripts уже ненужную
Reg Add HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\RunOnce /V RemoveFiles /D "cmd /C RD /S /Q C:\Windows\Setup\Scripts" /F

rem | Выключаем машину 
shutdown /s /t 0
