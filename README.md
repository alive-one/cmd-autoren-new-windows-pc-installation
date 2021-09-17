# cmd-autoren-new-windows-pc-installation
!!! Don't forget to change local administrator password (now it is qwerty-123) after deployment !!!

On every windows machine this script does:
01. Map specified network share as local disk Z: 
02. Takes Z:\inv-numbers.txt file with inventory numbers as input (see example of inv-numbers.txt here in main project folder). 
03. Rename current Windows machine according to inventory number from Z:\inv-numbers.txt file. (If inventory number is 010203 Windows system name will be I010203).
04. Write in file Z:\new-pc-name.txt string "mac-address pc-name" info (Ex. I010203 00-BB-CC-88-77-55).
05. Store hardware info for every system in html-file named as machine name (ex. I010203.html).
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
01. Create network share available so that Windows machines from deployed image could reach share and map using "net use" command. (You can change share name and path in script text).
02. In this share create file inv-numbers.txt and put there any list of names for new PCs in STRING divided by spaces, like 001 002 003 (See inv-numbers.txt here as example).
Be aware that in output file script will add capital letter "I" before theese names, so that PCs will be renamed like I001, I002, I003, etc. You can edit script to avoid it, just remember that Windows PC name can not contain digits only.
03. Also put in this folder file with Windows 10 and MS Office 13 key (See example w10-mso13-key.txt in this project files).

Part III. Deployng Image and getting results
01. Deploy captured in Part I image on desired amount of PCs using Clonezilla Lite Server or any mass deployment software of your choice.
(Make sure that you have internet connection for each PC, since without internet connection autounattend will not work in silent mode and ask for network setup).
02. Just restart macnines manually or set your software to autoreboot.
03. If everything were set up properly and with help of almighty Korgoth you end up with new-pc-names.txt file
where will be stored "pc-name mac-address" strings like "I001 00-03-7F-50-5C-0D"
04. Also for each PC you will have html-files like I010203.html with system name and main hardware info on your machines represented in form of simple tables readable by any browser.
