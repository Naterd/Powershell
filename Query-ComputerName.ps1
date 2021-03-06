<#
.SYNOPSIS
 	Query computer accounts with name of specified parameter, displays information in name,lastlogontimestamp format
.NOTES
    Author: Nathan Stewart
    Date:   7/9/2014 
#>


Param( 
	[Parameter(Mandatory=$true,HelpMessage='The computer name')]
	[string]$computername
	)
	
$computername = '*' + $computername + '*'
	
Get-ADComputer -Filter {name -like $computername} -properties name,lastlogondate | sort lastlogondate -Descending | ft name,lastlogondate