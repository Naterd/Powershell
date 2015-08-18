Param($CurrentUser,$NewUser)
if (!$CurrentUser) {$CurrentUser = Read-Host "User to copy groups from"}
if (!$NewUser) {$NewUser = Read-Host "New user to be added to groups"}
foreach ($user in $NewUser)
	{(Get-QADUser $CurrentUser).memberof | Add-QADGroupMember -Member amds\$NewUser}