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

$userobjects = Get-ADUser -Filter {enabled -eq "true"} -Properties name,title,department,manager,directreports,telephonenumber

$filename = 'C:\scripts\Advancedmd Associate Report - ' + (get-date -UFormat "%m-%d-%Y") + '.xlsx'

#Create Excel workbook object
$oXl = New-Object -ComObject "Excel.Application"
$oXl.Visible = $true
$oWb = $oXl.Workbooks.Add()
$oWs = $oWb.Worksheets.Add()
$oWs.Name = 'User Report'
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
	 $oCells.item($row,3) = $User.department 
     $oCells.item($row,4) = if ($user.Manager) {(get-ADuser ($User.Manager) -Properties Manager).Name }
     $oCells.item($row,5) = $reports
     $oCells.item($row,6) = $User.telephoneNumber	 
     $row++
}

AddWorksheetHeaders "Name" "Title" "Department" "Manager" "Direct Reports" "Telephone Number"





#Save workbook
$oWb.SaveAs($filename)
$oWb.Close()
$oXl.Quit()