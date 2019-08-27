REG ADD HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\W32Time\Parameters\ /v Type /t REG_SZ /d NTP
sleep 1
w32tm /config /manualpeerlist:"0.uk.pool.ntp.org,0x1 1.uk.pool.ntp.org,0x1 2.uk.pool.ntp.org,0x1 3.uk.pool.ntp.org,0x1"
sleep 1
w32tm /config /reliable:yes
sleep 1
net stop w32time && net start w32time
sleep 1
w32tm /resync /nowait