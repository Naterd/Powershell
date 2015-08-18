<#
.SYNOPSIS
	Scan Z drive and add AMDS\DLPFILERDR to each file/directory with Full Control
.NOTES
	Author: Nate Stewart with function from Joel Woodbury
	Date: 5/6/2014	
#>


$directories = get-childitem "##UPDATE##" -Force

$user1 = "##UPDATE##"
$path = $null

#Directories require different ACLs, this function differintiates between a file and a directory and sets permissions accordingly.
function Add-Permissions 
{
param (
	  [string]$path,
	  [string]$user, 
	  [string]$rights
	  )

    $isFolder = Test-Path $path -PathType Container
    
    if ($isFolder)
    {
        Write-host "Setting permissions for $path..." -foregroundcolor green
        $Acl = Get-Acl $path
        $rule = New-Object System.Security.AccessControl.FileSystemAccessRule("$($user)","$($rights)",'ContainerInherit, ObjectInherit', 'None', 'Allow')
        $Acl.AddAccessRule($rule)
        Set-Acl $path $Acl | Out-Null
    }
    else
    {
        Write-host "Setting permissions for $path..." -foregroundcolor green
        $Acl = Get-Acl $path
        $rule = New-Object System.Security.AccessControl.FileSystemAccessRule("$($user)","$($rights)",'Allow')
        $Acl.AddAccessRule($rule)
        Set-Acl $path $Acl | Out-Null
    }
}

foreach ($directory in $directories) {    
		
		 Add-Permissions $directory.FullName $user1 "FullControl"

			
			#Scan through directory for child items
            $files = Get-ChildItem ($directory.FullName) -Recurse | ForEach-Object {
             
                $filepath = $_.fullname
        
               	Add-Permissions $filepath $user1 "FullControl"
								
					  
         	 }
		  
	}
	

