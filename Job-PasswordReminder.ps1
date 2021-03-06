#Add the Quest PowerShell snapin
#written by Nate Stewart (desktop)
Add-PsSnapIn Quest.ActiveRoles.ADManagement -ErrorAction SilentlyContinue

$today = Get-Date
$logdate = Get-Date -format MM-dd-yy
$smtpserver = '##UPDATE##'
$emailFrom = '##UPDATE##'

$body ='
<h2>Please change your password to prevent loss of access to your systems.</h2>
<h2>If you are unable to change your password, please contact Desktop Support.</h2>
<text>How to change your password for local users (for remote users see paragraph 2)</text>
<ol>
<li>Begin by hitting Ctrl-Alt-Del and choosing the change password option.</li>
<li>Type your old password > new password and confirm it.</li>
<li>Click the arrow.</li>
</ol>
<h3>Additional step for remote users</h3>
<text>Because you are not at the corporate office, you <b>MUST</b> follow the steps below.</text>
<ol>
<li>Verify you are connected to the VPN through the Junos Pulse Client.</li>
<li>Follow steps listed above to change your password.</li>
</ol>
'



Get-QADUser -enabled -SizeLimit 0 | Select-Object samAccountName,mail,PasswordStatus | 
Where-Object {$_.PasswordStatus -ne 'Password never expires' -and $_.PasswordStatus -ne 'Expired' -and $_.PasswordStatus -ne 'User must change password at next logon.' -and $_.mail -ne $null} | 
ForEach-Object {

  $samaccountname = $_.samAccountName
  $mail = $_.mail 
  $passwordstatus = $_.PasswordStatus
  $passwordexpiry = $passwordstatus.Replace('Expires at: ','')
  $passwordexpirydate = Get-Date $passwordexpiry
  $daystoexpiry = ($passwordexpirydate - $today).Days
  
  if ($daystoexpiry -lt 5 ) {
    $emailTo = $mail
    $subject = "Your network account $samaccountname will expire in $daystoexpiry day(s) please change your password."    
    Send-MailMessage -To $mail -From $emailFrom -Subject $subject -Body $body -BodyAsHtml -SmtpServer $smtpserver
    #Write-Host "Email was sent to $mail, their password expires in $daystoexpiry day(s)"
    Add-Content \\sj-files\MP\IT\IT\PasswordReminder\$logdate.txt "Email was sent to $mail, their password expires in $daystoexpiry day(s)" -Force
  }
}
Send-MailMessage -To '##UPDATE##' -From '##UPDATE##' -Subject "Password expiring log for $today" -Body "This is the password expiring log from $today" -Attachments "\\sj-files\MP\IT\IT\PasswordReminder\$logdate.txt" -SmtpServer $smtpserver