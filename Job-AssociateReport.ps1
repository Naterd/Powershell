import-module ActiveDirectory

$excel = New-Object -ComObject "Excel.Application"
$excel.Visible =  $true
$workbook = $excel.Workbooks.Add()
$worksheet = $workbook.Worksheets.Add()
$worksheet = $workbook.Activesheet
$worksheet.name = "Associate Report"

$ocells = $worksheet.Cells

$filename = 'C:\scripts\' + 'Associate Report' + (Get-Date -Format M-dd-yy) + '.xlsx'

$users = Get-QADUser -Enabled -SizeLimit 5000 -DontUseDefaultIncludedProperties -IncludedProperties name,manager,directreports,department,office,creationdate,lastlogon | where {$_.office -ne $null} | sort name


Function AddWorksheetHeaders() {
    #This function adds the headers that are passed, then sets the column width.
    $col = 1
    $oColumns=$worksheet.Columns
    foreach ($arg in $args) {
          $ocells.item(1,$col)=$arg
          $ocells.item(1,$col).font.bold=$True
          $ocells.item(1,$col).font.underline=$True          
		  $oColumns.item($col).ColumnWidth = 30 	 
		 
          $col++
		  
		  }
		  
		  
	}
	
AddWorksheetHeaders "Name" "Manager" "Direct Reports" "Department" "Location" "Created" "LastLogon" "Title"

$row = 2

foreach ($user in $users) {	
	
	$ocells.item($row,1) = $user.name
	
	if ($user.Manager -ne $null) {
		
		#user.Manager returns the Distinguished name, its huge, using that value find that users full name
		$manager = Get-QADUser ($user.Manager) | select -expandproperty name
		$ocells.item($row,2) = $manager
		}
	
	if ($user.directreports -ne $null) { 	
		$reports = $user.directreports
		
		#instead of storing the distinguished name in the spreadsheet, find their full names and add it to a string
		foreach ($name in $reports) { 
			$fullname = Get-QADUser $name -DontUseDefaultIncludedProperties -Includedproperties name | select -ExpandProperty name
			$directreports += "$fullname, "
			}
			
		$ocells.item($row,3) = $directreports
			
		}	
	
	$ocells.item($row,4) = $user.department
	$ocells.item($row,5) = $user.office
	$ocells.item($row,6) = $user.creationdate
	$ocells.item($row,7) = $user.lastlogon
	$ocells.item($row,8) = $user.title
	$row++
	
	#null out variables for next user
	$manager = $null
	$directreports = $null
	
	
	}




#close out excel com objects
$workbook.SaveAs($filename)	
$workbook.Close()
$excel.Quit()

#There is a bug in powershell with excel com objects not releasing correctly and causing excel.exe process to not close
while( [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ocells)){}
while( [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet)){}
while( [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook)){}
while( [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)){}