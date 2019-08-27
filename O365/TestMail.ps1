$EmailFrom = read-host -prompt 'sending address' 
$EmailTo = read-host -prompt 'recipient address'
$Subject = “Test Message”
$Body = “This is just a test mail”
$SMTPServer = read-host -prompt 'email server address'
$SMTPClient = New-Object Net.Mail.SmtpClient($SmtpServer, 25)
$SMTPClient.Send($EmailFrom, $EmailTo, $Subject, $Body) 
