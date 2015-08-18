<#
.SYNOPSIS
	Pull mailboxes from OU and check for activesync devices and then remove them.
.EXAMPLE
	Remove-ActiveSyncDevices "##UPDATE##"
.NOTES
    Author: Nate Stewart
    Date: 4/10/2014
#>

Param(
[Parameter(Mandatory=$true,HelpMessage='Enter OU to remove activesync devices. IE ##UPDATE##')]
	[string]$OU
    )

$UserList = Get-CASMailbox -OrganizationalUnit $OU -Filter {hasactivesyncdevicepartnership -eq $true -and -not displayname -like "CAS_{*"} | Get-Mailbox

$UserList | foreach { 
                    Get-ActiveSyncDeviceStatistics -Mailbox $_ | remove-activesyncdevice -confirm:$false

                    $name = $_.name

                    Write-Host -ForegroundColor Yellow "Removing Devices for Mailbox $name"                    
                    
                    }

