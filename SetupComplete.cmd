rem ������� � ����� ���������� ����� �������� ��� ���� �������
del %SYSTEMDRIVE%\Unattend.xml

rem ������� ���� ������� ������� ��� ������ �������� ������ �������������� �������
Reg Add HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\RunOnce /V RenamePC /D "cmd /C C:\Windows\Setup\Scripts\autoren.bat" /F

