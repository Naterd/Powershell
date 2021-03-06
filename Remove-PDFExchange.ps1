﻿[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

$wmi = Get-WmiObject -Class win32_product

#Queries for installed pdf exchange versions
$pdf = $wmi | where {$_.Name -like "*PDF-XChange*"}

#for every incorrect version, removes it.
if ($pdf -ne $null)
{
	
	
	$prompt = [System.Windows.Forms.MessageBox]::Show("PDF exchange will now be removed and Adobe Reader installed in its place. To continue, all internet browsers will now be closed. Press OK when you are ready to proceed." , "CONTOSO" ,1,[System.Windows.Forms.MessageBoxIcon]::Warning,[System.Windows.Forms.MessageBoxDefaultButton]::Button1,[System.Windows.Forms.MessageBoxOptions]::DefaultDesktopOnly)
    
	if ($prompt -eq "OK") {
	
		foreach ($app in $pdf)
    	{
		
		#Get-Process iexplore* | kill
		#Get-Process chrome* | kill
		#Get-Process firefox* | kill
		
        &cmd /c "msiexec /x $($app.IdentifyingNumber) /passive"       
	    
    	}
	
		
		
				
		}
		
				
		
		
	}
#install adobe reader

$Arguments = @()
$Arguments += "/i"
$Arguments += "`"##UPDATE##\AdobeReader\AcroRead.msi`""
$Arguments += "/passive"
				
Start-Process "msiexec.exe" -ArgumentList $Arguments -Wait


