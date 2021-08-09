rem Удаляем с корня системного диска ненужный уже файл ответов
del %SYSTEMDRIVE%\Unattend.xml

rem Создаем ключ реестра который при логоне запустит скрипт переименования системы
Reg Add HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\RunOnce /V RenamePC /D "cmd /C C:\Windows\Setup\Scripts\autoren.bat" /F

