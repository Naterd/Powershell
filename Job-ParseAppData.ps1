<#
	.SYNOPSIS
		Scans through text files located in \\desktop01\Temp\scan\ and checks for non whitelisted files in the users appdata folder.
	.NOTES
		Author: Nate Stewart
		Date: 10/14/2013
	
#>

[array]$whitelist = Get-Content "\\desktop01\scripts\Job-ParseAppData\whitelist.txt"

$files = Get-content "\\desktop01\scripts\Test\acollett-10-14-13.txt"

foreach ($file in $files) {

$test =  $file | where { $_ -notmatch ([regex]::Escape($whitelist))}

$test


}

	
	

	
