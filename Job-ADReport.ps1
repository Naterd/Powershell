

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

$UserSearchDN = 'dc=CONTOSO,dc=CONTOSO,dc=com'
$ComputerSearchDN = 'OU=User Computers,DC=CONTOSO,DC=CONTOSO,DC=com'
$ServersSearchDN = 'OU=Servers,dc=CONTOSO,dc=CONTOSO,dc=com'
$DCSearchDN = 'OU=Domain Controllers,dc=CONTOSO,dc=CONTOSO,dc=com'

$userobjects = Get-ADUser -LDAPFilter "(name=*)" -SearchBase $UserSearchDN -Properties *

$sFileName = 'C:\scripts\ActiveDirectory Report - ' + (get-date -UFormat "%m-%d-%Y") + '.xlsx'
#$ping = new-object System.Net.NetworkInformation.Ping

#Create Excel workbook object
$oXl = New-Object -ComObject "Excel.Application"
$oXl.Visible = $true
$oWb = $oXl.Workbooks.Add()

#Add first worksheet & report members of admin groups
$oWs = $oWb.ActiveSheet
$oWs.name = 'Domain Admin Groups'
$oCells = $oWs.Cells
$GroupsToCheck = "Administrators","Domain Admins","Enterprise Admins","Backup Operators","Account Operators","Domain Controllers"
$row = 2
foreach ($grp in $GroupsToCheck) {
     $members = get-adgroupmember $grp |sort-object -property name
     foreach ($member in $members) {
          $oCells.item($row,1) = $grp
          $oCells.item($row,2) = $member.name
          $row++
     }
}
AddWorksheetHeaders "Group" "Member"

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
AddWorksheetHeaders "Name" "Email" "Expired" "Last Logon"

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
AddWorksheetHeaders "Name" "Email" "Password set"

#Add worksheet & report listing the domain accounts that have not logged in for 30 days or more
$oWs = $oWb.Worksheets.Add()
$oWs.Name = 'Inactive Users'
$oCells = $oWs.Cells
$Users = $userobjects | where-object { $_.LastLogonDate -lt (get-date).adddays(-180)} | sort-object -property 'name'
$row = 2
foreach ($User in $Users) {
     $oCells.item($row,1) = $User.name
     $oCells.item($row,2) = $User.mail
     $oCells.item($row,3) = $User.LastLogonDate
     $row++
}
AddWorksheetHeaders "Name" "Email" "Last Logon"

#Add worksheet & report listing the computer accounts that have not logged in for 30 days or more
$oWs = $oWb.Worksheets.Add()
$oWs.Name = 'Inactive Computers'
$oCells = $oWs.Cells
$Computers = Get-ADComputer -LDAPFilter "(name=*)" -SearchBase $ComputerSearchDN -Properties * | where-object { $_.LastLogonDate -lt (get-date).adddays(-30)} | sort-object -property 'name'
$row = 2

foreach ($Computer in $Computers) {
     $oCells.item($row,1) = $Computer.name
     $oCells.item($row,2) = $Computer.LastLogonDate     
     $row++
}
AddWorksheetHeaders "Computer" "Last Logon"


#Add worksheet of full report of AD
$oWs = $oWb.Worksheets.Add()
$oWs.Name = 'Full Report'
$oCells = $oWs.Cells
$Users = $userobjects | sort-object -property 'name'
$row = 2



foreach ($User in $Users) {
     #Check for direct reports, if they exist add them to $reports    
     $reports = $null   
     if ($user.directReports -ne $null) { ($user.directreports) | %{ $reports += (get-ADuser $_).name + ", " } }
     
     #populate spreadsheet   
     $oCells.item($row,1) = $User.name
     $oCells.item($row,2) = $User.title     
     $oCells.item($row,3) = if ($user.Manager) {(get-ADuser ($User.Manager) -Properties Manager).Name }
     $oCells.item($row,4) = $reports
     $oCells.item($row,5) = $User.CanonicalName
     $row++
}

AddWorksheetHeaders "Name" "Title" "Manager" "Direct Reports" "AD Path"





#Save workbook
$oWb.SaveAs($sFileName)
$oWb.Close()
$oXl.Quit()
