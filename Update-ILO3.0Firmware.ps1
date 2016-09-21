ynopsis
   Updates ilo3.0 firmware version to latest version supplied. Unless it is below 1.20, you have to update to 1.20 first before updating to latest version
.Description
   Keep in mind that the find-ilo command relies on the ilo OS responding to its query broadcast, I noticed depending on how the network
   was performing at that current time, some hosts would respond every time and some would respond off and on /shrugs
   In order to speed up the upgrade it uses the start-job command. Each upgrade is running in the background as a job. To check the status of the upgrade run Get-Job *
.EXAMPLE
   Update-ILO3.0Firmware -subnet 10.53.140.30-200 -firmware12path "PATH\1.20\ilo3_120.bin" -firmwarelatest "PATH\1.85\ilo3_185.bin" -firmwarelatestversion 1.85 
.NOTES   
   Author: Nate Stewart
   Date: 11-30-2015
#>

Param
    (        
        [Parameter(Mandatory=$true,Position=0,HelpMessage='Subnet range that command ips are on e.g 10.53.140.1-254')]
        [string]$subnet,
        [Parameter(Mandatory=$true,Position=1,HelpMessage='Full path to 1.2 firmware .bin file')]
        [string]$firmware12path,
        [Parameter(Mandatory=$true,Position=2,HelpMessage='Full path to latest version firmware .bin file')]
        [string]$firmwarelatest,
        [Parameter(Mandatory=$true,Position=2,HelpMessage='Lastest firmware version to update to. (What is listed in the name of .bin file e.g 2.20')]
        [double]$firmwarelatestversion

    )

Import-Module *HPilo* -ErrorAction Stop

$creds = Get-Credential -UserName administrator -Message 'HPIlo Admin Pass'

$ilo = Find-HPiLO $subnet

foreach ($ilohost in $ilo) {
        
        #zero out variables as precaution because of loop
        $firmwareversion = $null
        $ilohardware = $null
        $iloip = $null
        $ilogen = $null
              

        $firmwareversion = $ilohost.FWRI -as [double]
        $iloip = $ilohost.IP
        $ilogen = $ilohost.PN
        $ilohardware  = $ilohost.SPN
        
        #in order to not have to wait for each update one at a time we start them all as jobs
        if (($ilogen -like '*iLO 3*') -and ($firmwareversion -lt 1.2)) {           

            Start-Job -ArgumentList $iloip,$creds,$firmware12path -name $iloip -scriptblock {Update-HPiLOFirmware -Server $args[0] -Credential $args[1] -Location $args[2]}
            }

        if (($ilogen -like '*iLO 3*') -and ($firmwareversion -ge 1.2) -and ($firmwarelatestversion -gt $firmwareversion)) {            

            Start-Job -ArgumentList $iloip,$creds,$firmwarelatest -name $iloip -scriptblock {Update-HPiLOFirmware -Server $args[0] -Credential $args[1] -Location $args[2]}
            } 
        

                      
        $i++

        Write-Progress -activity "Starting Job for IP: $ilohp to update from version $firmwareversion" -status "Jobs started:$i " -PercentComplete (($i / $ilo.length)  * 100)

        $iloinfo = $null

}



