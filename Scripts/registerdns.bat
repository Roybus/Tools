ipconfig /flushdns
wait 1
ipconfig /registerdns
wait 1
net stop netlogon && net start netlogon