# cmd-autoren-new-windows-pc-installation
!!! Don't forget to change local administrator password (now it is qwerty-123) after deployment !!!

On every windows machine where it was launches this script does: 
00. Map network share \\ushare\backup\autoren\ as Z: disk (You can use \\your\own\share if you specify it in file).
01. Takes Z:\inv-numbers.txt file with inventory numbers as input (see example of inv-numbers.txt here in main project folder). 
02. Rename current Windows machine according to inventory number Z:\inv-numbers.txt file. (If inventory number is 010203 Windows system name will be I010203).
03. Write in file Z:\new-pc-name.txt string "mac-address pc-name" info (Ex. I010203 00-BB-CC-88-77-55).
04. Store hardware info for every system in html-file named as machine name (ex. I010203.html)
All in automode, no user interaction required. Better use this script for mass unattend deployment of windows OS.

Part I. Preparing Image
01. Install Windows 10 or Windows 7 and all necesary software.
02. Use sysprep to prepare your image for deployment
(Be aware that you need to remove user with name "user" in your image in audit sysprep mode since Unattend.xml stored in this repository create such user during installation process).
03. Reboot and load any Preinstallation Enviromnent of your choice (Strelec, HirenBCD or their Microsoft analog).
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
01. Deploy captured in Part I image on needed amount of PCs using Clonezilla Lite Server or any mass deployment software of your choice.
(Make sure that you have internet connection for each PC, since autounattend will not work in silent mode and ask for network setup).
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
