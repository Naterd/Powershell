Param($UserName)

if (!$UserName) {$UserName = Read-Host "User to export mailbox from"}

$Today=get-date -uformat %d-%b-%Y

$Path="\\desktop01\Files\Archived E-mail\"

$CompletedPATH=$Path+($UserName)+($Today).ToString()+".pst"

New-MailboxExportRequest -Mailbox $UserName -FilePath $CompletedPATH -BadItemLimit 500 -AcceptLargeDataLoss