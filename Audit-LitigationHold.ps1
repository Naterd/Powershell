<#
.Synopsis
   Scan for mailboxes with litigation enabled, report back litigation hold stats and total storage space in litigation.
   It will export to a csv if supplied a path
.EXAMPLE
   .\Audit-LitigationHold.ps1
.EXAMPLE
   .\Audit-LitigationHold.ps1 -ExportCSV C:\scripts\report.csv
.PARAMETER ExportCSV
    Path to export the csv to
.NOTES
   Date: 9-18-15
   Author: Nate Stewart
#>

Param(
[Parameter(Position=1)]
[string]$ExportCSV
)


$mailboxes = get-mailbox -Filter { LitigationHoldEnabled -eq 'True'}

foreach ($mailbox in $mailboxes) {

    $totalsize = $mailbox | Get-MailboxStatistics | Select-Object -ExpandProperty TotalDeletedItemSize

    #create new powershell object that includes the totalsize property
    $usermailbox = New-Object PSObject -Property @{
        
        Name           = $mailbox.Name
        Enabled        = $mailbox.LitigationHoldEnabled
        Date           = $mailbox.LitigationHoldDate
        LitigationMail = $totalsize

    }

    [array]$report += $usermailbox

}

#sort largest mailboxes to top
$report = $report | Sort-Object LitigationMail -Descending

if ($ExportCSV) { $report | Export-Csv -Path $ExportCSV -Force }

if (!$ExportCSV) { $report | Format-Table Name,LitigationMail,Enabled,Date }