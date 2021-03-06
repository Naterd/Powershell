<#
.SYNOPSIS
 	Disables all user accounts located in text file, format text file as one username per line.
.PARAMETER Filepath
    Path to text file containing usernames.
.EXAMPLE
    Disable-UserList C:\scripts\names.txt
.NOTES
    Author: Nathan Stewart
    Date:   1/14/2014  
#>

Param(
[Parameter(Mandatory=$True,Position=1)]
[string]$filepath 
)

$users = Get-Content $filepath

foreach ($user in $users)  {

	Write-Host "Disabling $user"

	#Disable user account
	Disable-QADUser $user | Out-Null
	
	Remove-QADMemberOf -Identity $user -RemoveAll	
		
	#set description with termination date
	
	$date = Get-Date -Format "dd MMMM yyyy"
	
	Set-QADUser $user -Description "Terminated: $date" | Out-Null	
	
	Write-Host "Hidding account:$user from exchange lists"
	
	set-mailbox $user -HiddenFromAddressListsEnabled $true
	
	#forward mail to user specified
	
	$forwardmail = Read-Host "Would you like to forward their mail to another user? (y/n)"
	
	if ($forwardmail -like "y") {
	
		try {
			$mailbox = get-mailbox $user.SamAccountName
			$forwarduser = Read-Host "Please enter the account name of the user to recieve the forwarded emails"
			$forwarduser = Get-QADUser $forwarduser -WildcardMode PowerShell
		
			#if the supplied user exists forward their mail
			if ($forwarduser) {
			
			$forwardname = $forwarduser.samaccountname
			$mailbox | set-mailbox  -forwardingaddress $forwardname -delivertomailboxandforward $true -HiddenFromAddressListsEnabled $true
			
			Write-Host $forwardname "`nhas been set as forwarding address"
			
			}
			
			}
			
		catch {
		
		Write-Host "`nAn error has occurred please rerun script" -ForegroundColor Green
		
		}
	
	}
	}	
