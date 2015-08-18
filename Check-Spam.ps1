$Server = "SERVERNAME"
$logs = get-messagetrackinglog -Server $Server -Start ((Get-Date).AddHours(-12)) -End (get-date) -ResultSize unlimited
Write-Host -ForegroundColor Green "$Server top senders in last 12 hours"
$logs | ?{$_.Sender -notlike "*.int"} | group Sender | sort Count -desc | select -First 20 | ft Count, Name -auto
Write-Host -ForegroundColor Green "$Server top subject lines in last 12 hours"
$logs | ?{$_.Sender -notlike "*.int"} | group MessageSubject | sort Count -desc | select -First 20 | ft Count, Name -auto
    