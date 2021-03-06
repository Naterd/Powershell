<#
	.SYNOPSIS
		Scan Active Directory Computer Objects and move them to OldComputers OU to be disabled in the future
	.DESCRIPTION
		If the lastlogontimestamp is greater than 6 months the computer is inactive and should be removed from the domain
	.EXAMPLE
		This script runs as a job, to run manually simple run the script .\Job-DisableOldComputers.ps1	
	.NOTES
		Author: Nathan Stewart

#>

#Query current date
$currentdate = get-date

#days to move inactive computer accounts (6 months)
$cutoff = $currentdate.AddDays(-180)

#Query computers from AD
$computers = Get-QADComputer -IncludedProperties lastlogontimestamp -SearchRoot "##UPDATE##"

#Check each computers logontimestamp
foreach ($computer in $computers) {
		
		#if the last logon of the computer to the domain is greater than 6 months add the object to the movecomputers variable
		if (($computer.lastlogontimestamp) -lt $cutoff) {
		
		#add computer to array of computer objects to be moved at later point
		[array]$movecomputers += $computer 
	
		}
		
		
	}

#Move the objects
$movecomputers | Move-QADObject -NewParentContainer "##UPDATE##"


#HTML Header for notification email
$a = @"
<style>
BODY{font-family: Verdana, Arial, Helvetica, sans-serif;font-size:12;font-color: #000000}
TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}
TH{border-width: 2px;padding: 0px;border-style: solid;border-color: black;background-color: #d01b1b}
TD{border-width: 2px;padding: 3px;border-style: solid;border-color: black;background-color: #ffffff}
</style>
"@

#body of email in html
$body += "<text><b>The following computer(s) were moved to the OldComputers OU due to inactivity:</text>"
$body += $movecomputers | select name | sort -Property name | ConvertTo-Html -Head $a


$subject = "Active Directory Inactive Computer Audit: $currentdate"

$smtpserver = "##UPDATE##"			
			
Send-MailMessage -To "##UPDATE##" -From "##UPDATE##" -Subject $subject -Body $body -BodyAsHtml -SmtpServer $smtpserver