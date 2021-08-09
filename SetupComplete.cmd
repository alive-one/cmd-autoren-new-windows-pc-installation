rem | Remove no longer neccessary answers file
del %SYSTEMDRIVE%\Unattend.xml

rem | Create regedit key which will start autoren.bat on logon
Reg Add HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\RunOnce /V RenamePC /D "cmd /C C:\Windows\Setup\Scripts\autoren.bat" /F
