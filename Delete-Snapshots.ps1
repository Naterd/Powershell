foreach ($guest in Get-Cluster VDI | get-vm)
{
	$snaps = Get-Snapshot -VM $guest
	foreach ($snap in $snaps)
	{
		if ($snap -ne $null)
		{
		Remove-Snapshot -Snapshot $snap -RemoveChildren -Confirm:$false
		}
	}
}