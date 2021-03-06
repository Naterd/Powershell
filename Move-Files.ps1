<#
.SYNOPSIS
	Import files from .txt to be moved to secure location \\desktop01\files\archived user data\move-files	
.DESCRIPTION	
	Attempts to move file, if an error occurrs it stops immediately and notifies	
.NOTES		
	Author: Nathan Stewart - Desktop Team
#>
Param(
    [Parameter(mandatory=$true,HelpMessage='Enter the path to text file containing filenames to be moved')]
    [string]$files
    )

$list = get-content $files
$targetlocation ='\\desktop01\files\Archived User Data\move-files'
$logdate = Get-Date -format MM-dd-yy
$smtpserver = "##UPDATE##"
$emailFrom = "##UPDATE##"


foreach ($file in $list) {

		if (Test-Path $file) {
	
		#grab filename and owner
		$fullname = (Get-Item $file).fullname
		$shortname = (Get-Item $file).name
		$owner = (Get-Acl $fullname).Owner

        
	
		    Try {
		        #attempt to move item, if it fails stop immediately
		        Move-Item -Path "$fullname" -Destination $targetlocation -ErrorAction Stop -ErrorVariable $errormove -Force
		        }
	
		    Catch {
		
		        #Error occurred
	            Write-Host -ForegroundColor Yellow "$shortname failed to move"
                $errormove
		        }

            #build list of files moved
            $filesmoved = New-Object PSObject
		    $filesmoved | Add-Member NoteProperty -Name FileName -Value $shortname		
		    $filesmoved | Add-Member NoteProperty -Name Owner -Value $owner
		    [array]$totalfiles += $filesmoved

            $errormove = $null


    }
		
		
	
}

if ($totalfiles) {	
	
	
	[string]$body = $totalfiles | ConvertTo-Html
	
	Send-MailMessage -To "##UPDATE##" -From $emailFrom -SmtpServer $smtpserver  -Subject "The following files were moved from the Z drive on $logdate" -Body $body -BodyAsHtml
    
    
	}
	

