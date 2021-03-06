<#
.SYNOPSIS
 	Create new user AD account based on information pulled from Jira
.PARAMETER casenumber
    This is the number it will use to pull the information. If the new associate request was ITSD-2935 you would enter only 2935
.EXAMPLE
    New-AssociateFromJira 2935
.NOTES
    Author: Nathan Stewart
    Date:   4/2/2015   
#>


Param(
[Parameter(Mandatory=$True,Position=1)]
[string]$casenumber
)

[String] $uri = '##UPDATE##/jira/rest/api/2/issue/ITSD-' + $casenumber;

#The only way to get invoke-restmethod to pre authenticate so that jira will work
#$base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(('{0}:{1}' -f $username,$password)))
$casedata = Invoke-restmethod -Headers @{Authorization=('Basic {0}' -f '##UPDATE##')} -Uri $uri -Method Get

#pull new associate request info from jira
$name = $casedata.fields.customfield_10907
$department = $casedata.fields.customfield_10908
$title = $casedata.fields.customfield_10303
$mirrorassociate = $casedata.fields.customfield_11013.name
$officephone = $casedata.fields.customfield_11511
$manager = $casedata.fields.customfield_11010.name
$office = $casedata.fields.customfield_11510.value

#Derive username
[array]$usernamesplit = $name -split '\s+'

#Checks for the availability of the username, if it is taken it counts the similar usernames and increments by +1
#Trim first and last name to username format

$fname = $usernamesplit[0].ToLower()
$fname = $fname.Substring(0,1)
$lname = $usernamesplit[1].ToLower()
$username = $fname + $lname

#tests if the username exists, if yes count them and increment by 1
if((get-aduser -filter {samAccountname -eq $username}) -ne $null) { 

	[string[]]$users = Get-aduser -filter "samaccountname -like '$username*'" | Select-Object -ExpandProperty samAccountname
	
	$username = $username + ($users.count)
	
	}


function QueryMailStoreSize {

	#This functions finds the smallest mailstore to be used in creating the new user	
	#filter out the default mailbox database before sorting

	Get-MailboxDatabase -Status | Where-Object -FilterScript { $_.name -ne 'Public Folder'} | Sort-Object databasesize | Select-Object -first 1	}
	
function HomeDriveOffice {
	
	#based upon location, sets the home drive for either mbs or south jordan
	$mbsdrive = "##UPDATE##\$username"
	$sjdrive = "##UPDATE##\$username"
	$path = $sjdrive

    switch ($office)  {    
    
	    'Springfield' { Set-ADUser $username -HomeDrive 'P:' -HomeDirectory $mbsdrive -Office $office | Out-Null; $path = $mbsdrive; $OU ='##UPDATE##/SPRINGFIELD'  }
		#because it is a MBS associate, change the path to their homefolder to the local storage in Springfield		
		
	    'South Jordan'{ Set-ADUser $username -HomeDrive 'P:' -HomeDirectory $sjdrive -Office $office; $OU ='##UPDATE##/SOUTHJORDAN' | Out-Null }
		
        'Remote' { Set-ADUser $username -HomeDrive 'P:' -HomeDirectory $sjdrive -Office $office; $OU ='##UPDATE##/SOUTHJORDAN' | Out-Null }		
		
	    'India' { Set-ADUser $username -HomeDrive 'P:' -HomeDirectory $sjdrive -Office $office; $OU ='##UPDATE##/INDIA' | Out-Null }

        'Manila' { $OU ='##UPDATE##/INDIA' }
		
		default { Set-ADUser $username -HomeDrive 'P:' -HomeDirectory $sjdrive -Office 'South Jordan'; $OU ='##UPDATE##/SOUTHJORDAN' | Out-Null }

        }
		
	try {
		#Create homedrive folder and set permissions	
		
		$user = "AMDS\$username"
		$InheritanceFlag = [System.Security.AccessControl.InheritanceFlags]::ContainerInherit -bor [System.Security.AccessControl.InheritanceFlags]::ObjectInherit
		$PropagationFlag = [System.Security.AccessControl.PropagationFlags]::None
		$objType = [System.Security.AccessControl.AccessControlType]::Allow 

		New-Item $path -type directory
		$acl = Get-Acl $path
		$permission = $user,'Modify', $InheritanceFlag, $PropagationFlag, $objType
		$accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule $permission
		$acl.SetAccessRule($accessRule)
		Set-Acl $path -AclObject $acl
		
	}
	
	Catch {
		
		Write-Host '`nThere was an error creating the home drive folder. Please verify you are running this script as Domain Admin' -ForegroundColor green
		Write-Host $error
		
		
		#prompt user if they want to remove the incorrectly created user account
		$ask = Read-Host "Would you like to delete the account: $username"
			if ($ask -like 'y') {
				Remove-ADObject $username
				Break;
			}
			else {
				Break;
			}
	
		}
	}	


$principalname = $username + '##UPDATE##'

#find smallest mailstore
$mailstore = QueryMailStoreSize

Write-Host "`nName: $name Username: $username Maildatabase: $mailstore"
Write-Host "`nDepartment: $department Title: $title"  
Write-Host "`nOffice Phone: $officephone Manager: $manager Office: $office"

Write-Host "`nCreating Mailbox"
New-Mailbox -Name $name -DisplayName $name -UserPrincipalName $principalname -Alias $username -OrganizationalUnit NEWUSERS -Database $mailstore -Password (ConvertTo-SecureString -String '##UPDATE##' -AsPlainText -Force) -FirstName ($usernamesplit[0]) -LastName ($usernamesplit[1]) | Out-Null


Write-Host "`nWaiting for AD Replication"

#this prevents issues with replication not happening fast enough and the commands erroring out, not ideal but it works.
Start-Sleep 9

set-aduser $username -Department $department -Title $title -OfficePhone $officephone -Manager $manager 

. HomeDriveOffice	

#mirror permissions
Write-host "`nMirroring AD groups from: $mirrorassociate"

$groups = (Get-ADUser $mirrorassociate -Properties MemberOf).MemberOf

$groups | Add-ADGroupMember -Members $username

#Enable Lync

Try {
    Enable-CsUser $username -RegistrarPool '##UPDATE##' -SipAddress "sip:$username@##UPDATE##"

    Write-Host "`nLync has been enabled"
    }

Catch {
    Write-Host 'An error has occurred, do you have the lync tools installed? Do you have permission to create lync users?' -ForegroundColor Green

    Write-Host "`nLync has not been enabled, please enable Lync after the completion of this script"

}




Write-Host "`nPosting update in Jira"

$groupnames = $groups | ForEach-Object { get-adgroup $_ | Select-Object -ExpandProperty name}

$date = Get-date

#Silly formatting to make it cleaner
$groupnames = $groupnames | % { $_ + ', '}

#contains escaped jira text formatting
$body = @{
    body = "`{color:red`}Account was created: $date`{color`}" + "`nAdded user to following AD groups" + "`n`{quote`}$groupnames`{quote`}"
    visibility = @{
        type = 'role'
        value = 'Service Desk Team'
        }
        
}

$body = $body | ConvertTo-Json

$uri = '##UPDATE##/jira/rest/api/2/issue/ITSD-' + $casenumber + '/comment'

#Post groups copied and creation notice in Jira
Invoke-restmethod -Headers @{Authorization=('Basic {0}' -f '##UPDATE##')} -Uri $uri -Method POST -body $body -ContentType 'application/json' | Out-Null

	




  
