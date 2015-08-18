$email = Read-Host 'Please enter the email string (wildcard usage is allowed, even ENCOURAGED.)'
Write-Host "You searched for $email" -foregroundcolor yellow
get-qadobject -LdapFilter "(proxyaddresses=$email)" | Format-Table Name,Type,Email