net stop wuauserv && net stop bits && net stop dosvc
wait 99999
net start wuauserv && net start bits && net start dosvc