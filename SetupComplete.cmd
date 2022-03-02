rem | Get rid of password expires issue, for password change alert will stop script execution
wmic path Win32_UserAccount where (name="user") SET PasswordExpires=FALSE
 
rem | Delete Unattend.xml whic is no longer neccessary
del %SYSTEMDRIVE%\Unattend.xml
 
rem | Create Regedit key which set system to copy and execute autoren.bat after autologon
Reg Add HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\RunOnce /V RenamePC /D "cmd /C C:\Windows\Setup\Scripts\cp-autoren.bat" /F
