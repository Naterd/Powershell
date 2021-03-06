

#Job-ADReport.ps1
#Creates an Excel spreadsheet with reports on information from AD

#Reference http://www.petri.co.il/export-to-excel-with-powershell.htm
#Note: When referencing cells, it's (Row,Column)

import-module ActiveDirectory

Function AddWorksheetHeaders() {
    #This function adds the headers that are passed, then autofits the column width
    $col = 1
    $oColumns=$oWs.Columns
    foreach ($arg in $args) {
          $oCells.item(1,$col)=$arg
          $oCells.item(1,$col).font.bold=$True
          $oCells.item(1,$col).font.underline=$True
          $oColumns.item($col).EntireColumn.Autofit()
          $col++
     }
}

$UserSearchDN = 'ENTERDNHERE'
$ComputerSearchDN = 'ENTERDNHERE'
$ServersSearchDN = 'ENTERDNHERE'
$DCSearchDN = 'ENTERDNHERE'

$userobjects = Get-ADUser -filter {enabled -eq $true} -SearchBase $UserSearchDN -Properties name,lastlogondate,mail,passwordexpired,cannotchangepassword,passwordneverexpires,passwordlastset

$sFileName = 'C:\scripts\ActiveDirectory Security Report - ' + (get-date -UFormat '%m-%d-%Y') + '.xlsx'

#Create Excel workbook object
$oXl = New-Object -ComObject 'Excel.Application'
$oXl.Visible = $true
$oWb = $oXl.Workbooks.Add()

#Add first worksheet & report members of admin groups
$oWs = $oWb.ActiveSheet
$oWs.name = 'Domain Admin Groups'
$oCells = $oWs.Cells
$GroupsToCheck = 'Administrators','Domain Admins','Enterprise Admins','Backup Operators','Account Operators','Domain Controllers'
$row = 2
foreach ($grp in $GroupsToCheck) {
     $members = get-adgroupmember $grp |sort-object -property name
     foreach ($member in $members) {
          $oCells.item($row,1) = $grp
          $oCells.item($row,2) = $member.name
          $row++
     }
}
AddWorksheetHeaders 'Group' 'Member'

#Add worksheet & report listing the domain accounts with expired passwords
$oWs = $oWb.Worksheets.Add()
$oWs.Name = 'Expired Passwords'
$oCells = $oWs.Cells
$Users = $userobjects | where-object { $_.PasswordExpired -and !$_.CannotChangePassword -and !$_.PasswordNeverExpires} | sort-object -property 'name'
$row = 2
foreach ($User in $Users) {
     $oCells.item($row,1) = $User.name
     $oCells.item($row,2) = $User.mail
     if ($User.PasswordLastSet) { $oCells.item($row,3) = (Get-Date -Date $User.PasswordLastSet).AddDays(90) }
     $oCells.item($row,4) = $User.LastLogonDate
     $row++
}
AddWorksheetHeaders 'Name' 'Email' 'Expired' 'Last Logon'

#Add worksheet & report listing the domain accounts with passwords that do not expire
$oWs = $oWb.Worksheets.Add()
$oWs.Name = 'Passwords do not Expire'
$oCells = $oWs.Cells
$Users = $userobjects | where-object { $_.PasswordNeverExpires} | sort-object -property 'name'
$row = 2
foreach ($User in $Users) {
     $oCells.item($row,1) = $User.name
     $oCells.item($row,2) = $User.mail
     $oCells.item($row,3) = $User.PasswordLastSet
     $row++
}
AddWorksheetHeaders 'Name' 'Email' 'Password set'

#Add worksheet & report listing the domain accounts that have not logged in for 60 days or more
$oWs = $oWb.Worksheets.Add()
$oWs.Name = 'Inactive Users'
$oCells = $oWs.Cells
$Users = $userobjects | where-object { $_.LastLogonDate -lt (get-date).adddays(-60)} | sort-object -property 'name'
$row = 2
foreach ($User in $Users) {
     $oCells.item($row,1) = $User.name
     $oCells.item($row,2) = $User.mail
     $oCells.item($row,3) = $User.LastLogonDate
     $row++
}
AddWorksheetHeaders 'Name' 'Email' 'Last Logon'

#Add worksheet & report listing the computer accounts that have not logged in for 60 days or more
$oWs = $oWb.Worksheets.Add()
$oWs.Name = 'Inactive Computers'
$oCells = $oWs.Cells
$Computers = Get-ADComputer -filter {enabled -eq $true} -SearchBase $ComputerSearchDN -Properties lastlogondate,name | where-object { $_.LastLogonDate -lt (get-date).adddays(-60)} | sort-object -property 'name'
$row = 2

foreach ($Computer in $Computers) {
     $oCells.item($row,1) = $Computer.name
     $oCells.item($row,2) = $Computer.LastLogonDate     
     $row++
}

AddWorksheetHeaders 'Name' 'Last Logon'

#Save workbook
$oWb.SaveAs($sFileName)
$oWb.Close()
$oXl.Quit()
