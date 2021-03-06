Add-PSSnapin Quest.ActiveRoles.ADManagement -ErrorAction SilentlyContinue

$arrayComp ="##UPDATE##" 
$now = (get-date)
$smtpserver = "##UPDATE##"

$a = "<style>"
$a = $a + "BODY{font-family: Verdana, Arial, Helvetica, sans-serif;font-size:12;font-color: #000000}"
$a = $a + "TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}"
$a = $a + "TH{border-width: 2px;padding: 0px;border-style: solid;border-color: black;background-color: #d01b1b}"
$a = $a + "TD{border-width: 2px;padding: 3px;border-style: solid;border-color: black;background-color: #ffffff}"
$a = $a + "</style>"

foreach ($machine in $arrayComp)

{  
	$PrintJobs = Get-WmiObject Win32_PrintJob -computername $machine | where-object {$_.dmtfDate -ne 1}



if($PrintJobs -ne $NULL){ 

	foreach ($PrintJob in $PrintJobs) { 
	
	$then = [System.Management.ManagementDateTimeConverter]::ToDateTime($printjob.timesubmitted)
	$age = $now - $then 
	
		if($age.minutes -gt 15) { 
			
			
			$name = $PrintJob.Name.Split(",")
			$owner = $PrintJob.Owner
			$ownername = Get-QADUser "$owner" | Select-Object -ExpandProperty Name
			$document = $PrintJob.Document
			
			$printer = Get-WmiObject Win32_Printer -ComputerName $machine | where {$_.name -eq $name[0]}
			$printername = $printer.Name
			$printer.CancelAllJobs()			
			
			$subject =$ownername + "'s printing job has been cancelled on " + $printername		

			$body = $printer | Select-Object name,location,systemname | ConvertTo-Html -Head $a
			$body += "<br><text>User: <b>$ownername<b> was trying to print $document</text>"
			
			Send-MailMessage -To "##UPDATE##" -From "##UPDATE##" -Subject $subject -Body ($body | Out-String) -BodyAsHtml -SmtpServer $smtpserver  
			Send-MailMessage -To "$owner@##UPDATE##" -From "##UPDATE##" -Subject "$ownername your $document has been cancelled on $printername" -Body "This email is to notify you the document you are trying to print has stalled and was auto cancelled for impeding the printer queue." -SmtpServer $smtpserver
			
	
			}
		}
	}
}