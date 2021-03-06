<#
	.SYNOPSIS
	 Scan user computer directory for infection files (c:\users). If found dump a text file to temp directory with the full path to the file
	
	.EXAMPLE
	 .\Job-CheckInfection.ps1
	 
	.NOTES
	 Date: 10/22/2104
	 Author: Nate Stewart		

#>
$date = Get-Date -Format MM-dd-yy
$path = "\\desktop01\temp\scan\$env:COMPUTERNAME-$env:USERNAME-$date.txt"

#Files to search for
$include = @("obupdat.*","INSTALL_TOR.*","DECRYPT_INSTRUCTION.*")

$badfiles = Get-childitem -Recurse "C:\Users\*" -Include $include | select -ExpandProperty fullname

if (!(Test-Path $path) -and ($badfiles)) {
Add-Content $badfiles -Path $path -Force

}
