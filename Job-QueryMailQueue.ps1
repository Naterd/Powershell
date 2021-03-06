<#
.SYNOPSIS
	Query Exchange Outbound Mail Queue. If the messages waiting to be sent exceed 10.
	Turn off the TPSI send connector and set the send connector back to MXLogic SmartHost.
.NOTES
	Author: Nate Stewart	
#>

$now = (get-date)

$smtpserver = "##UPDATE##"

$messages = get-queue sj-exch-fe\10026 | select -ExpandProperty messagecount

$mxenabled = get-sendconnector -identity "MXLogic SmartHost" | select -ExpandProperty enabled
$tpsienabled = get-sendconnector -identity "TPSI Interceptor" | select -ExpandProperty enabled


if (($messages -ge 10) -and ($mxenabled -eq "False")) {
	
	#Get-SendConnector | Set-SendConnector -Enabled $false	
	#Set-SendConnector -identity "MXLogic SmartHost" -Enabled $true	
	Send-MailMessage -To "##UPDATE##" -From "##UPDATE##" -Subject "TPSI send connector backlogged, MXlogic send connector re-enabled - $now" -SmtpServer $smtpserver
	
	}
	
	
if (($messages -eq 0) -and ($tpsienabled -eq "False")) {
	
	#Get-SendConnector | Set-SendConnector -Enabled $false
	#Set-SendConnector -identity "TPSI Interceptor" -Enabled $true	
	Send-MailMessage -To "##UPDATE##" -From "##UPDATE##" -Subject "Mail queue backlogged cleared, reverting back to TPSI send connector - $now" -SmtpServer $smtpserver
	
	}