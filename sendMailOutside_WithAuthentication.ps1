#By vmt1991
#Tool used for send mail with authentication with Exchange account to outside mail
#More tool at My Blog: http://vmt1991.pythonanywhere.com/



$SmtpServer = "cas.domain.com"

$SmtpPort = "25"
 
# Message stuff

$MessageSubject = "Mail test"

$Message = New-Object System.Net.Mail.MailMessage $MailFrom,$MailTo

$Message.IsBodyHTML = $true

$Message.Subject = $MessageSubject
 
[void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
 
$title = 'Send Mail Outside Tool v1.0 - By vmt1991: '
 
$msg   = 'Subject Email Send:'
 
$subject_send = [Microsoft.VisualBasic.Interaction]::InputBox($msg, $title)
 
$msg   = 'Input Email Send:'
 

$email_send = [Microsoft.VisualBasic.Interaction]::InputBox($msg, $title)

 

$msg   = 'Input Password:'

$password_send = [Microsoft.VisualBasic.Interaction]::InputBox($msg, $title)

 
$msg   = 'List mail customer: '

$user_lst = [Microsoft.VisualBasic.Interaction]::InputBox($msg, $title)


$msg   = 'HTML content mail: '

$content = [Microsoft.VisualBasic.Interaction]::InputBox($msg, $title)

$Username = $email_send.Split('@')[0]

Write-Host "Email Send: $Username"

Write-Host "Password: $password_send"

Write-Host "Userlist: $user_lst"

Write-Host "Content send: $content"

$content_mail=Get-Content $content

$Smtp = New-Object Net.Mail.SmtpClient($SmtpServer,$SmtpPort)
 
$Smtp.EnableSsl = $true

ForEach ($line in (Get-Content $user_lst))
{

Write-Host "----------Send mail to user $line-------------------"

$Smtp.Credentials = New-Object System.Net.NetworkCredential($Username,$Password)

$Message = New-Object System.Net.Mail.MailMessage $email_send,$line

$Message.IsBodyHTML = $true

$Message.Subject = $subject_send

$Message.Body = $content_mail

$Smtp.Send($Message)
}