
:: ##################################################
::
:: Script Name:  reset_WindowsUpdate.bat
:: By:  Zack Thompson / Created:  6/22/2016
:: Version:  1.0 / Updated:  6/22/2016 / By:  ZT
:: Description:  This script resets the Windows Update Components.
::		See:  https://support.microsoft.com/en-us/kb/971058
::
:: ##################################################

:: Stop the BITS service, the Windows Update service, and the Cryptographic service
net stop bits
net stop wuauserv
net stop appidsvc
net stop cryptsvc

:: Delete the qmgr*.dat files
Del "%ALLUSERSPROFILE%\Application Data\Microsoft\Network\Downloader\qmgr*.dat"

:: Rename the softare distribution folders backup copies
Ren %systemroot%\SoftwareDistribution SoftwareDistribution.bak
Ren %systemroot%\system32\catroot2 catroot2.bak


:: Reset the BITS service and the Windows Update service to the default security descriptor
:: Uncomment the following two lines for "aggressive mode".
:: sc.exe sdset bits D:(A;;CCLCSWRPWPDTLOCRRC;;;SY)(A;;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;BA)(A;;CCLCSWLOCRRC;;;AU)(A;;CCLCSWRPWPDTLOCRRC;;;PU)
:: sc.exe sdset wuauserv D:(A;;CCLCSWRPWPDTLOCRRC;;;SY)(A;;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;BA)(A;;CCLCSWLOCRRC;;;AU)(A;;CCLCSWRPWPDTLOCRRC;;;PU)


:: Reregister the BITS files and the Windows Update files
cd /d %windir%\system32
regsvr32.exe atl.dll
regsvr32.exe urlmon.dll
regsvr32.exe mshtml.dll
regsvr32.exe shdocvw.dll
regsvr32.exe browseui.dll
regsvr32.exe jscript.dll
regsvr32.exe vbscript.dll
regsvr32.exe scrrun.dll
regsvr32.exe msxml.dll
regsvr32.exe msxml3.dll
regsvr32.exe msxml6.dll
regsvr32.exe actxprxy.dll
regsvr32.exe softpub.dll
regsvr32.exe wintrust.dll
regsvr32.exe dssenh.dll
regsvr32.exe rsaenh.dll
regsvr32.exe gpkcsp.dll
regsvr32.exe sccbase.dll
regsvr32.exe slbcsp.dll
regsvr32.exe cryptdlg.dll
regsvr32.exe oleaut32.dll
regsvr32.exe ole32.dll
regsvr32.exe shell32.dll
regsvr32.exe initpki.dll
regsvr32.exe wuapi.dll
regsvr32.exe wuaueng.dll
regsvr32.exe wuaueng1.dll
regsvr32.exe wucltui.dll
regsvr32.exe wups.dll
regsvr32.exe wups2.dll
regsvr32.exe wuweb.dll
regsvr32.exe qmgr.dll
regsvr32.exe qmgrprxy.dll
regsvr32.exe wucltux.dll
regsvr32.exe muweb.dll
regsvr32.exe wuwebv.dll

:: Reset Winsock and Proxy
netsh winsock reset
netsh winhttp reset proxy

:: Restart the BITS service, the Windows Update service, and the Cryptographic service
net start bits
net start wuauserv
net start appidsvc
net start cryptsvc