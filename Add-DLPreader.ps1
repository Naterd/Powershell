Param( 
	[Parameter(Mandatory=$true,HelpMessage='The mount point to add perms to. EX: Z:\MP\Backups')]
	[string]$path
)

$user = "SJ-FILES01\dlpfilerdr"
$InheritanceFlag = [System.Security.AccessControl.InheritanceFlags]::ContainerInherit -bor [System.Security.AccessControl.InheritanceFlags]::ObjectInherit
$PropagationFlag = [System.Security.AccessControl.PropagationFlags]::None
$objType = [System.Security.AccessControl.AccessControlType]::Allow 

$directories = get-childitem $path -force
$directories | foreach {
    $acl = Get-Acl $_.FullName
	$permission = $user,"Read","Allow" #$InheritanceFlag, $PropagationFlag, $objType
	$accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule $permission
	$acl.SetAccessRule($accessRule)
	Set-Acl -Path $_.FullName -AclObject $acl
    }

    