
net stop "System Event Notification Service" /y
 
net stop "Background Intelligent Transfer Service" /y
 
net stop "COM+ Event System" /y
 
net stop "Microsoft Software Shadow Copy Provider" /y
 
net stop "Volume Shadow Copy" /y
 
cd /d %windir%\system32 
 
net stop vss 
 
net stop swprv 
 
regsvr32 /s ATL.DLL
 
regsvr32 /s comsvcs.DLL
 
regsvr32 /s credui.DLL
 
regsvr32 /s CRYPTNET.DLL
 
regsvr32 /s CRYPTUI.DLL
 
regsvr32 /s dhcpqec.DLL
 
regsvr32 /s dssenh.DLL
 
regsvr32 /s eapqec.DLL
 
regsvr32 /s esscli.DLL
 
regsvr32 /s FastProx.DLL
 
regsvr32 /s FirewallAPI.DLL
 
regsvr32 /s kmsvc.DLL
 
regsvr32 /s lsmproxy.DLL
 
regsvr32 /s MSCTF.DLL
 
regsvr32 /s msi.DLL
 
regsvr32 /s msxml3.DLL
 
regsvr32 /s ncprov.DLL
 
regsvr32 /s ole32.DLL
 
regsvr32 /s OLEACC.DLL
 
regsvr32 /s OLEAUT32.DLL
 
regsvr32 /s PROPSYS.DLL
 
regsvr32 /s QAgent.DLL
 
regsvr32 /s qagentrt.DLL
 
regsvr32 /s QUtil.DLL
 
regsvr32 /s raschap.DLL
 
regsvr32 /s RASQEC.DLL
 
regsvr32 /s rastls.DLL
 
regsvr32 /s repdrvfs.DLL
 
regsvr32 /s RPCRT4.DLL
 
regsvr32 /s rsaenh.DLL
 
regsvr32 /s SHELL32.DLL
 
regsvr32 /s shsvcs.DLL
 
regsvr32 /s /i swprv.DLL
 
regsvr32 /s tschannel.DLL
 
regsvr32 /s USERENV.DLL
 
regsvr32 /s vss_ps.DLL
 
regsvr32 /s wbemcons.DLL
 
regsvr32 /s wbemcore.DLL
 
regsvr32 /s wbemess.DLL
 
regsvr32 /s wbemsvc.DLL
 
regsvr32 /s WINHTTP.DLL
 
regsvr32 /s WINTRUST.DLL
 
regsvr32 /s wmiprvsd.DLL
 
regsvr32 /s wmisvc.DLL
 
regsvr32 /s wmiutils.DLL
 
regsvr32 /s wuaueng.DLL
 
sfc /SCANFILE=%windir%\system32\catsrv.DLL
 
sfc /SCANFILE=%windir%\system32\catsrvut.DLL
 
sfc /SCANFILE=%windir%\system32\CLBCatQ.DLL
 
Takeown /f %windir%\winsxs\temp\PendingRenames /a 
icacls %windir%\winsxs\temp\PendingRenames /grant "NT AUTHORITY\SYSTEM:(RX)"
icacls %windir%\winsxs\temp\PendingRenames /grant "NT Service\trustedinstaller:(F)"
icacls %windir%\winsxs\temp\PendingRenames /grant BUILTIN\Users:(RX)
Takeown /f %windir%\winsxs\filemaps\* /a 
icacls %windir%\winsxs\filemaps\*.* /grant "NT AUTHORITY\SYSTEM:(RX)"
icacls %windir%\winsxs\filemaps\*.* /grant "NT Service\trustedinstaller:(F)"
icacls %windir%\winsxs\filemaps\*.* /grant BUILTIN\Users:(RX)

net stop cryptsvc
net start cryptsvc

net start "COM+ Event System"

net start "dhcp server"

net start "file replication service"

net start "System Event Notification Service"
 
net start "Background Intelligent Transfer Service"
 
net start "Microsoft Software Shadow Copy Provider"
 

