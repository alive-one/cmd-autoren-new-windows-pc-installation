# cmd-autoren-new-windows-pc-installation
Take txt-file with inventory numbers as input, rename windows machine and store hardware info in html-file and "mac-address pc-name" info in txt-file on any available specified net share. All in automode, no user interaction required. Better use this script for mass unattend deployment of windows OS.

Part I. Preparing Image
01. Install Windows 10 or Windows 7 and all necesary software.
02. Use sysprep to prepare your image for deployment
(Be aware that you need to remove user with name "user" in your image, in audit sysprep mode, since Unattend.xml stored in this repository create such user during installation process).
03. Reboot and load any Preinstallation Enviromnent of your choice (Strelec, HirenBCD or their Microsoft preinstallation analog).
04. Put Unattend.xml in root of system drive with your Windows OS (Usually it is disk C:)
05. Put autoren.bat and SetupComplete.cmd in Windows/Setup/Scripts (If "Scripts" folder does not exist, make one).
06. Reboot, boot from USB or network and load any Imaging Software of your choice (Clonezilla, Acronis, DISM, etc.) 
07. Create disk image of your yet undeployed Microsoft OS and store it anywhere you like (Local Disk, Network Share, etc.)

Part II. Preparing Network Share
01. Create network share \\your-server\backup\autoren available via local net, or anyway so that Windows machine from your freshly deployed image could reach it and map using "net use" command. (You can change name of share and path in script if you wish, of course).
02. In this share create file \\your-server\backup\autoren\inv-numbers.txt and put there any list of names for new PCs (For example 001, 002, 003 and so on).
(Be aware that in output file script will add capital letter "I" before theese names, so that PCs will be renamed like I001, I002, I003, etc. You can edit script to avoid it, just remember that Windows PC name can not contain digits only).
03. Also put in this folder file with Windows 10 and MS Office 13 key (see example in this project files) and name it \\your-server\backup\autoren\inv-numbers.txt\w10-mso13-key.txt

Part III. Deployng Image and getting results
01. Deploy captured in PartI image on needed amount of PCs using Clonezilla Lite Server or any mass deployment software of your choice.
02. Just restart macnines manually or set your software to autoreboot.
03. If everything were set up properly and with help of almighty Korgoth you end up with \\your-server\backup\autoren\new-pc-names.txt file
where will be stored "pc-name mac-address" strings like 
I001 00-03-7F-50-5C-0D
I002 00-03-7F-60-7C-0D
I003 00-03-7F-70-6C-0D
04. Also for each PC you will have html-files
\\your-server\backup\autoren\I001.html
\\your-server\backup\autoren\I002.html
\\your-server\backup\autoren\I003.html
with hardware info on your machines, represented in form of simple tables.

P.S. Be wary that there is no known way to get proper amount of Graphical adapter memory above 4Gb, using AdapterRAM command. 
P.P.S But there are rumors that it is still possible. So, there is room for improovment. (c) Count Dooku
