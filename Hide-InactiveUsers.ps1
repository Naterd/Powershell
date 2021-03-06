Import-Csv C:\scripts\inactivefull.csv | ForEach-Object {
$first = $_.firstname
$last = $_.lastname
$name = "$first " + "$last"

$user = Get-QADUser -disabled $name

	if($user) {
		
		$mailbox = get-mailbox $user.SamAccountName
		
		$mailbox.name
		
		$mailbox | set-mailbox -HiddenFromAddressListsEnabled $true
	
		}
$user = $null
}
