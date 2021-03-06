<#
.SYNOPSIS
 	Disables specified user account and removes group memberships in AD. 
	Will assign a forwarding address for all future mail and hide user from the Global Address List if you use -ForwardMail.
    To post a log to the jira case, use -UpdateInJira casenumber
.PARAMETER UserAccount
    This is the user account to be disabled in samAccountname form. Current naming conventions are first initial last name.
	Nate Stewart would be nstewart. Beware of users with similar account names, if you arent sure double check before running script.
.PARAMETER UpdateInJira
    Supply this parameter and it will attempt to post termination notice to the jira case, you need to supply the case number for it to do this
    If the case is ITST-2944 you would enter 2944
.PARAMETER ForwardMail
    Username of mailbox that the terminated user's mail will be sent to. Please use SamAccountName, Nathan Stewart would be nstewart, be sure to verify
.EXAMPLE
    disable-user ttest
.EXAMPLE 
    disable-user ttest 2987
.EXAMPLE
    disable-user -useraccount ttest -updateinjira 2987 -forwardmail nstewart
.NOTES
    Author: Nathan Stewart
    Date:   4/9/2015   
#>

Param(
[Parameter(Mandatory=$True,Position=1)]
[string]$UserAccount,

[Parameter(Position=2)]
[string]$UpdateInJira,

[Parameter(Position=3)]
[string]$ForwardMail
)

$user = Get-Aduser $UserAccount

if ($user)  {

	#Disable user account
	Disable-ADAccount $user | Out-Null
	
    #remove group memberships
	$groups = (get-aduser $user -Properties memberof).memberof
    $groups | Remove-ADGroupMember -Members $user -Confirm:$false	
		
	#set description with termination date
	
	$date = Get-Date	
	Set-ADUser $user -Description "Terminated: $date" -Office $null -Manager $null -OfficePhone $null | Out-Null	
	
	Write-Host 'Hidding account from Global Address List'
	
	set-mailbox $user.samAccountname -HiddenFromAddressListsEnabled $true	
	
	if ($forwardmail) {
	
		try {
			$mailbox = get-mailbox $user.SamAccountName			
			$forwarduser = Get-ADUser $ForwardMail
		
			#if the supplied user exists forward their mail
			if ($forwarduser) {
			
			$forwardname = $forwarduser.samaccountname
			$mailbox | set-mailbox  -forwardingaddress $forwardname -delivertomailboxandforward $true
			
			Write-Host $forwardname 'has been set as forwarding address'
			
			}
			
			}
			
		catch {
		
		Write-Host "`nAn error has occurred please verify the user set as the forward has a mailbox" -ForegroundColor Green
		
		}
	
	}


    if($UpdateInJira) {

        $groupnames = $groups | ForEach-Object { get-adgroup $_ | Select-Object -ExpandProperty name}

        #Silly formatting to make it cleaner
        $groupnames = $groupnames | % { $_ + ', '}

        #contains escaped jira text formatting
        $body = @{
            body = "`{color:red`}Account terminated: $date`{color`}" + "`nRemoved user from following AD groups" + "`n`{quote`}$groupnames`{quote`}"
            visibility = @{
                type = 'role'
                value = 'Service Desk Team'
            }
        
        }

        $body = $body | ConvertTo-Json

        $uri = 'http://sj-jira:8081/jira/rest/api/2/issue/ITSD-' + $UpdateInJira + '/comment'

        #Post groups copied and creation notice in Jira
        Invoke-restmethod -Headers @{Authorization=('Basic {0}' -f 'REDACTED')} -Uri $uri -Method POST -body $body -ContentType 'application/json' | Out-Null
    }
}
	
else {

	Write-Host "`nThe user:" $UserAccount "was not found. Please verify the correct username is being entered`n" -ForegroundColor Green
	
	}




	
































