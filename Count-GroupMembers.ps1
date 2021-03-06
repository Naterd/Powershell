<#
	.SYNOPSIS
		Count how many users are a part of a security group including nested groups
	.NOTES
		Author: Nate Stewart
		Date: 1/8/15
#>

[CmdletBinding()]
Param( 
	[Parameter(Mandatory=$true,HelpMessage='Group to be counted')]
	[string]$group
	)
	
$global:count = 0
$ErrorActionPreference = "Stop"

function groupmember ($group) {

$groups = Get-ADGroupMember $group | ForEach-Object { 
	if($_.objectclass -eq "group") {	
	groupmember $_
	}
	else { 
	$global:count++	
	
	}
	}
	

}

groupmember $group

Write-Host $count