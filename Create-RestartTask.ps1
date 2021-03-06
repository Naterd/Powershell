<#
	.SYNOPSIS
		Create Scheduled Restart Job on host
	.DESCRIPTION
		Creates schedule task that restarts computer DAILY at 9PM
	.AUTHOR
		Nathan Stewart

#>

$action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument '-NoProfile -WindowStyle Hidden -command "& Restart-Computer -force"'

$trigger =  New-ScheduledTaskTrigger -Daily -At 9pm

Register-ScheduledTask -Action $action -Trigger $trigger -TaskName "Restart" -Description "Daily computer restart"
