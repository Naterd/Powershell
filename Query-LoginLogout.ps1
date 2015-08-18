<#
.SYNOPSIS
 	Query user login and logout event logs
.EXAMPLE
    Query-LoginLogout
.NOTES
    Author: Nathan Stewart
    Date:   7/2/2013   
#>

$logon = Get-EventLog "Security" |where { ($_.instanceid -eq 4648) -and (($_.timewritten -gt(get-date 7/1/13)))}

$lock =  Get-EventLog "Security" |where { ($_.instanceid -eq 4800) -and (($_.timewritten -gt(get-date 7/1/13)))} 