Param([Parameter(Mandatory = $True,
ValueFromPipeLine = $False,
Position = 0)]
[Alias('')]
[String]$ComputerName = "localhost"
)
$LastBoot = (Get-WmiObject -Class Win32_OperatingSystem -computername $computername).LastBootUpTime
$sysuptime =[System.Management.ManagementDateTimeconverter]::ToDateTime($LastBoot)

Write-Host "Last boot time for $computername was" $sysuptime

