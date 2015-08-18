$Date = Get-Date -uformat "%Y%m%d" 
$Logfile = "C:\Logs\ActiveSync-all-$date.txt" 
$Devices = @()

Add-Content -path $LogFile "name,devicemodel,devicetype,useragent,lastsynctime"


$Mailboxes = Get-CASMailbox -ResultSize Unlimited | Where {$_.HasActiveSyncDevicePartnership -eq $True -and $_.ExchangeVersion.ExchangeBuild.Major -ilike "14"}

ForEach ($mailbox in $mailboxes){ 
    $Devices= Get-ActiveSyncDeviceStatistics -Mailbox $mailbox.name 
    $name = $mailbox.Name 
    ForEach ($device in $devices) { 
        $Model = $Device.DeviceModel 
        $Type = $Device.DeviceType 
        $LastSyncTime = $Device.LastSuccessSync 
        $UserAgent = $Device.DeviceUserAgent 
        Add-Content -path $Logfile "$name,$Model,$Type,$UserAgent,$LastSyncTime" 
    } 
}