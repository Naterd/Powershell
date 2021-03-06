<#
.SYNOPSIS
	Scan Exchange enviroment for activesynce devices active in the last 30 days and export them to a spreadsheet
.DESCRIPTION		
	Exports the Name, Device Model, Device Type, Device OS and Last sync time.
.AUTHOR
    Nate Stewart		
#>


#create excel object
$excel = New-Object -ComObject "Excel.Application"
$excel.Visible =  $true
$workbook = $excel.Workbooks.Add()
$worksheet = $workbook.Worksheets.Add()
$worksheet = $workbook.Activesheet
$worksheet.name = "ActivesyncDevicesReport"
$ocells = $worksheet.Cells

$filename = 'C:\scripts\' + 'ActivesyncReport-' + (Get-Date -Format M-dd-yy) + '.xlsx'


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
	
AddWorksheetHeaders "Name" "DeviceModel" "DeviceType" "DeviceOS" "LastSyncTime"

$row = 2

$Mailboxes = Get-CASMailbox -ResultSize Unlimited | Where {$_.HasActiveSyncDevicePartnership -eq $True}

$Devices = @()

#Populate spreadsheet

foreach ($mailbox in $mailboxes) {

            $Devices= Get-ActiveSyncDeviceStatistics -Mailbox $mailbox.name | where { $_.LastSuccessSync -gt ((Get-Date).AddDays(-30)) }
            
            ForEach ($device in $devices) {
            
            
                $ocells.item($row,1) = $mailbox.name
                $ocells.item($row,2) = $Device.DeviceModel            
                $ocells.item($row,3) = $Device.DeviceType
                $ocells.item($row,4) = $Device.DeviceOS 
                $ocells.item($row,5) = $Device.LastSuccessSync
                
                $row++                
        } 
        
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

Write-Host "`n""Saving file to $filename $directory"