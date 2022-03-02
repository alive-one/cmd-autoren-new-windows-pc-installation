rem | Fix disabled anonimous connection to samba-server (Later we will restore this back to default)
reg add HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\LanmanWorkstation\Parameters /v AllowInsecureGuestAuth /t reg_dword /d 00000001 /f

rem | Fix disabled anonimous connection to samba-server (Later we will restore this back to default)
reg add HKLM\Software\Policies\Microsoft\Windows\LanmanWorkstation /v AllowInsecureGuestAuth /t reg_dword /d 00000001 /f

rem | Фиксим ебанатскую ошибку 85 которую не могут уже много лет пофиксить мудаки и пидорасы из Микрософт блядского
reg add "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager" /t REG_DWORD /v ProtectionMode /d 0x00000000 /f

rem | Монтируем сетевой диск 
rem | лучше по IP потому что DNS у винды бывает глючит
net use Z: \\192.168.0.70\backup\autoren

rem | Копируем с примонтированного диска файл autoren.bat
rem | и запускаем его
copy /y z:\autoren.bat C:\Windows\Setup\Scripts\autoren.bat
call C:\Windows\Setup\Scripts\autoren.bat
