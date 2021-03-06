<#	
	.DESCRIPTION
		Import user details from csv and import the values into Active Directory
	.EXAMPLE
		.\Update-UserInformation $PATHTOCSV	
	.NOTES
		Author: Nathan Stewart

#>

[CmdletBinding()]
Param( 
	[Parameter(Mandatory=$true,HelpMessage='Filepath of CSV')]
	[string]$CSV
	)
	
$userinformation = Import-Csv $CSV

foreach ($user in $userinformation) {
	
	#must use LDAP names of each field not the powershell names when using -replace in set-aduser
	#if entry in spreadsheet is empty, ignore it
	
	$username = $user.Name
	$usersam = Get-ADUser -Filter {name -eq $username} | select -ExpandProperty samAccountName
	
	$hash = @{}
  	if(!($user.department -eq "")){$hash.Department = $user.department;}
 	if(!($user.title -eq "")){$hash.Title = $user.title;}
 	if(!($user.telephoneNumber -eq "")){$hash.telephoneNumber = $user.telephoneNumber;}		
	
	#Check if hash is empty	
	if($hash.Count -gt 0) {
	
		Write-Host "Updating user: $username"
		Set-ADUser $usersam -Replace $hash 
		
		}
		
		#The manager property requires more logic to update
		if(!($user.manager -eq "")){	
			
			$managername = $user.Manager
			$manager = Get-ADUser -Filter {name -eq $managername} | select -ExpandProperty SamAccountName
			Set-ADUser $usersam -Manager $manager
		
		}
			
				
	}
		
	
	
	


	
	