#Count the mailboxes on each database in Exchange and report it to the shell
#Author: Nate Stewart
#Date: 6-18-14

$databases = Get-MailboxDatabase -status | select name,databasesize

$count = 0

foreach($database in $databases) {

	Write-Progress -Activity "Querying Databases" -Status "$database" -CurrentOperation "$count completed"

	$number = (get-mailbox -database $database.name).count
	
	$object = New-Object PSObject -Property @{
	
		Name = $database.name
		Mailboxes = $number
		DatabaseSize = $database.databasesize		
		
		}
	
	[array]$objectlist += $object
	
	$count++
	
	}
	
	$objectlist | ft