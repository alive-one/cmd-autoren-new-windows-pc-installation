rem | Fix disabled anonimous connection to samba-server (Later we will restore this back to default)
reg add HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\LanmanWorkstation\Parameters /v AllowInsecureGuestAuth /t reg_dword /d 00000001 /f

rem | Fix disabled anonimous connection to samba-server (Later we will restore this back to default)
reg add HKLM\Software\Policies\Microsoft\Windows\LanmanWorkstation /v AllowInsecureGuestAuth /t reg_dword /d 00000001 /f

rem | Fix "Error 85" issue (Later we will restore it back to default).
reg add "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager" /t REG_DWORD /v ProtectionMode /d 0x00000000 /f

rem | Mount network share 
net use Z: \\192.168.0.70\backup\autoren

rem | Copy from mounted share file autoren.bat 
copy /y z:\autoren.bat C:\Windows\Setup\Scripts\autoren.bat

rem | And execute it
call C:\Windows\Setup\Scripts\autoren.bat
