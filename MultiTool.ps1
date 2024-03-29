Write-Host '1 - Change calendar permissions - O365.'
Write-Host '2 - VSS Fix.'
Write-Host '3 - Change Users Password. - O365 User'
Write-Host '4 - DISM Repair.'
Write-Host '5 - SFC Scan.'
Write-Host '6 - Change UPN O365.'
Write-Host '7 - Stop Windows Updates - Period of time.'
Write-Host '8 - Register DNS - Server.'
Write-Host '9 - Config to NTP.'
Write-Host '10 - Display Email Alias for all Users - O365/Exchange.'
Write-Host '11 - Mass Import Users to O365 Using CSV File.'
Write-Host '12 - Grant Admin Full Access to all Mailboxes.'
Write-Host '13 - Grant a User Access to Another Mailbox.'
Write-Host '14 - Change all O365 password'

Write-Host "`n`n0 - Close."

#Function for selection table
function selction {
    
    $Number = Read-Host 'Please enter a number: 1-14' #Get user selection

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
        changeUserPasswordO365
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
    elseif ($Number -eq 11) {
        importO365CSV
    }
    elseif ($Number -eq 12){
        grantAdminMailbox
    }
    elseif ($Number -eq 13){
        mailboxAccess
    }
    elseif ($Number -eq 14){
        passO365Bulk
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
    Start-Sleep -S 2 #Sleep for 2 seconds

    #Connect to 0365 service
    $LiveCred = Get-Credential
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $LiveCred -Authentication Basic -AllowRedirection
    Import-PSSession $Session
    
    $ID = Read-Host 'Please enter the email for the account you wish to share the calendar from' #Get ID of user you are chnaging the permissions on
    $User = Read-Host 'Please enter the user you wish to have access to this' #Get ID of user you are giving access to
    $Access = Read-Host 'Please enter the level of access e.g. Editor' #Get the level of access

    Set-MailboxFolderPermissions -Identity $ID -User $User+':\Calendar' -AccessRights $Access

    Remove-PSSession $Session
    $LiveCred = ''

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

#Function to change user password from UPN
function changeUserPasswordO365 {
    
    #Connect to 0365 service
    $LiveCred = Get-Credential
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $LiveCred -Authentication Basic -AllowRedirection
    Import-PSSession $Session
    
    $o365User = Read-Host 'Please enter the username'
    $newPass = Read-Host 'Please enter the new password'
    $forceChange = Read-Host 'Would you like to force the user to change on next login? 1 - Yes | 2 - No'


    if ($forceChange -eq 1){
        Set-MsolUserPassword -UserPrincipalName $o365User -NewPassword $newPass -ForceChangePassword $True
    }
    else {
        Set-MsolUserPassword -UserPrincipalName $o365User -NewPassword $newPass -ForceChangePassword $False
    }
    
    Remove-PSSession $Session
    $LiveCred = ''

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
    
    #Connect to 0365 service
    $LiveCred = Get-Credential
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $LiveCred -Authentication Basic -AllowRedirection
    Import-PSSession $Session

    $currentUPN = Read-Host 'Please enter the current email for the user account.' #Get current username of account
    $newUPN = Read-Host 'Please enter the new email.' #Get new username for account

    Set-MsolUserPrincipalName -UserPrincipalName $currentUPN -NewUserPrincipalName $newUPN

    Remove-PSSession $Session
    $LiveCred = ''

    anyInput
   
}

#Funtion to stop Windows updates for a period of time
function stopWinUp {
    
    #Get Wait timer
    $timer = Read-Host 'Please enter a time (in seconds) for the wait period. 1-99999'
    
    net stop wuauserv
    net stop bits
    net stop dosvc
    Start-Sleep -S $timer
    net start wuauserv
    net start bits
    net start dosvc

    anyInput
}

#Function to register DNS on server
function registerDNS {
    ipconfig /flushdns
    Start-Sleep -S 1
    ipconfig /registerdns
    Start-Sleep -S 1
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

#Function for displaying all users O365
function displayEmailAlias {
    
    #Connect to O365 admin using PS session
    $LiveCred = Get-Credential
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $LiveCred -Authentication Basic -AllowRedirection
    Import-PSSession $Session

    Get-Mailbox | Select-Object DisplayName,@{Name=”EmailAddresses”;Expression={$_.EmailAddresses |Where-Object {$_ -LIKE “SMTP:*”}}} | Sort | Export-Csv C:\Users\%Username%\Desktop\email-aliases.csv

    Write-Host 'The list has been exported to your Desktop - email-aliases.csv'

    Remove-PSSession $Session
    $LiveCred = ''

    anyInput

}

#Fucntion for mass importing users into O365
function importO365CSV {
    
    #Connect to O365 admin using PS session
    $LiveCred = Get-Credential
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $LiveCred -Authentication Basic -AllowRedirection
    Import-PSSession $Session

    $pathCSV = Read-Host 'Please enter the path for the CSV file - Including name'

    Import-Csv $pathCSV | ForEach {
        New-MsolUser -UserPrincipalName $_.UserPrincipalName -DisplayName $_.DisplayName -FirstName $_.Firstname -Lastname $_.Lastname -Password $_.Password -UsageLocation $_.Location -ForceChangePassword
    }
    
    Remove-PSSession $Session
    $LiveCred = ''

    anyInput
}

#Function to grant the admin user full access to all mailboxes
function grantAdminMailbox {
    
    #Connect to O365 admin using PS session
    $LiveCred = Get-Credential
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $LiveCred -Authentication Basic -AllowRedirection
    Import-PSSession $Session

    $adminUser = Read-Host 'Please enter the admin user email.'
    Get-Mailbox -ResultSize unlimited -Filter {(RecipientTypeDetails -eq 'UserMailbox') -and (Alias -ne 'Admin')} | Add-MailboxPermission -User $adminUser -AccessRights FullAccess -InheritanceType All
    
    Remove-PSSession $Session
    $LiveCred = ''

    anyInput
}

#Function to grant user access to another mailbox O365
function mailboxAccess {
    
    #Connect to O365 admin using PS session
    $LiveCred = Get-Credential
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $LiveCred -Authentication Basic -AllowRedirection
    Import-PSSession $Session
    
    $idGrant = Read-Host 'Please enter the email of the mailbox you are giving access to'
    $userGrant = Read-Host 'Please enter the email of the user you wish to grant access to'
    $accessGrant = Read-Host 'Please enter the level of access e.g. FullAccess'
    $autoMap = Read-Host 'Would you like this to map on Outlook? 1 - Yes | 2 - No'

if ($autoMap -eq 1){
    Add-MailboxFolderPermission -Identity $idGrant -User $userGrant -AccessRights $accessGrant -AutoMapping:$true
}
else{
    Add-MailboxFolderPermission -Identity $idGrant -User $userGrant -AccessRights $accessGrant -AutoMapping:$false
}
    
    Remove-PSSession $Session
    $LiveCred = ''

    anyInput
}

#Function for change all O365 users' password
function passO365Bulk {
    
    #Connect to 0365 service
    $LiveCred = Get-Credential
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $LiveCred -Authentication Basic -AllowRedirection
    Import-PSSession $Session
    
    $newPass = Read-Host 'Please enter the new password'
    $forceChange = Read-Host 'Would you like to force the user to change on next login? 1 - Yes | 2 - No'


    if ($forceChange -eq 1){
        Get-MsolUser |%{Set-MsolUserPassword -userPrincipalName $_.UserPrincipalName –NewPassword $newPass -ForceChangePassword $True}
    }
    else {
        Get-MsolUser |%{Set-MsolUserPassword -userPrincipalName $_.UserPrincipalName –NewPassword $newPass -ForceChangePassword $False}
    }

    Remove-PSSession $Session
    $LiveCred = ''

    anyInput
    
}

#Import Needed modules
Import-Module MSOnline

selction