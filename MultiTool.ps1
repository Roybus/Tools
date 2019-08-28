Write-Host '1 - Change calendar permissions.'
Write-Host '2 - VSS Fix.'
Write-Host '3 - Change Users Password.'
Write-Host '4 - DISM Repair.'
Write-Host '5 - SFC Scan.'
Write-Host '6 - Change UPN.'
Write-Host '7 - Stop Windows Updates - Period of time'
Write-Host '8 - Register DNS - Server'
Write-Host '9 - Config to NTP'
Write-Host '10 - Display Email Alias for all Users - O365/Exchange'

Write-Host "`n`n0 - Close."

#Function for selection table
function selction {
    
    $Number = Read-Host 'Please enter a number: 1-10' #Get user selection

    if ($Number -eq 0){
        exit
    }
    elseif ($Number -eq 1){
        calPermissions
    }
    elseif ($Number -eq 2){
        vssFix
    }
    elseif ($Number -eq 3){
        changeUserPassword
    }
    elseif ($Number -eq 4){
        dism
    }
    elseif ($Number -eq 5){
        sfc
    }
    elseif ($Number -eq 6){
        upn0365
    }
    elseif ($Number -eq 7) {
        stopWinUp
    }
    elseif ($Number -eq 8) {
        registerDNS
    }
    elseif ($Number -eq 9) {
        registerNTP
    }
    elseif ($Number -eq 10) {
        displayEmailAlias
    }
    else {
        Write-Host 'Please enter a valid number.'
    }
    
}

#Function for asking user if they want to choose another option or close program
function anyInput {
    $option = Read-Host 'Select 1 to go back to selection or anything else to close...'
    
    if ($option -eq 1){
        selection
    }
    else{
        exit
    }
}

#Function for changing calendar permissions
function calPermissions {
    
    Write-Host 'Tool used to configure a users calendar permissions.'
    Start-Sleep -S 2 #Sleep for 2 secondsa
    Connect-MsolService #Connect to 0365 service
    $ID = Read-Host 'Please ennter the email for the user account.' #Get ID of user you are chnaging the permissions on
    $User = Read-Host 'Please enter the user you wish to have access e.g. Default' #Get ID of user you are giving access to
    $Access = Read-Host 'Please enter the level of access e.g. FullAccess' #Get the level of access

    Set-MailboxFolderPermissions -Identity $ID -User $User -AccessRights $Access

    anyInput
}

#Function to reset all VSS drivers
function vssFix {
        
        #Stop running services
        net stop "System Event Notification Service" /y
        net stop "Background Intelligent Transfer Service" /y
        net stop "COM+ Event System" /y
        net stop "Microsoft Software Shadow Copy Provider" /y
        net stop "Volume Shadow Copy" /y
    
        #Change to the windows DIR
        cd /d %windir%\system32 

        #Stop VSS Serivces
        net stop vss
        net stop swprv 

        #Take ownership of regkeys
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

        #Restart crypt service
        net stop cryptsvc
        net start cryptsvc
    
        #Start all services back up
        net start "COM+ Event System"
        net start "dhcp server"
        net start "file replication service"
        net start "System Event Notification Service"
        net start "Background Intelligent Transfer Service"
        net start "Microsoft Software Shadow Copy Provider"
    }

#Function to change user password from UPN - Working Progress
function changeUserPassword {
    $UPN = $_.UserPrincipalName; Get-ADUser -Filter {UserPrincipalName -Eq $UPN} - Properties *
    
    anyInput
}

#Function for DISM repair
function dism {
    dism /online /cleanup-image /restorehealth
}

#Function for running SFC
function sfc {
    sfc /scannow
}

#Function to change UPN in O365
function upn0365 {
    
    Write-Host 'Tool used to configure a users Username for O365.'
    Start-Sleep -S 2 #Sleep for 2 secondsa
    Connect-MsolService #Connect to 0365 service
    $currentUPN = Read-Host 'Please ennter the current email for the user account.' #Get current username of account
    $newUPN = Read-Host 'Please enter the new email.' #Get new username for account

    Set-MsolUserPrincipalName -UserPrincipalName $currentUPN -NewUserPrincipalName $newUPN -AccessRights $Access

    anyInput
   
}

#Funtion to stop Windows updates for a period of time
function stopWinUp {
    
    #Get Wait timer
    $timer = Read-Host 'Please enter a time (in seconds) for the wait period. 1-99999'
    
    net stop wuauserv && net stop bits && net stop dosvc
    wait $timer
    net start wuauserv && net start bits && net start dosvc

    anyInput
}

#Function to register DNS on server
function registerDNS {
    ipconfig /flushdns
    wait 1
    ipconfig /registerdns
    wait 1
    net stop netlogon
    net start netlogon
    
    anyInput
}

#Function to register DC to NTP UK time
function registerNTP {
    REG ADD HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\W32Time\Parameters\ /v Type /t REG_SZ /d NTP
    start.sleep 1
    w32tm /config /manualpeerlist:"0.uk.pool.ntp.org,0x1 1.uk.pool.ntp.org,0x1 2.uk.pool.ntp.org,0x1 3.uk.pool.ntp.org,0x1"
    start.sleep 1
    w32tm /config /reliable:yes
    start.sleep 1
    net stop w32time 
    net start w32time
    start.sleep 1
    w32tm /resync /nowait
    
    anyInput
}

#Function for displaying all users mailboxes
function displayEmailAlias {
    get-mailbox | foreach{ 

        $host.UI.Write("White", $host.UI.RawUI.BackGroundColor, "`nUser Name: " + $_.DisplayName+"`n")
       
        for ($i=0;$i -lt $_.EmailAddresses.Count; $i++){
           $address = $_.EmailAddresses[$i]
           
           $host.UI.Write("White", $host.UI.RawUI.BackGroundColor, $address.AddressString.ToString()+"`t")
        
           if ($address.IsPrimaryAddress)
           { 
               $host.UI.Write("Green", $host.UI.RawUI.BackGroundColor, "Primary Email Address`n")
           }
          else
          {
               $host.UI.Write("Green", $host.UI.RawUI.BackGroundColor, "Alias`n")
           }
        }
    }
    
    anyInput

}

selction