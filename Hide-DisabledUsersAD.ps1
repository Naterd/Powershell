$users = get-content C:\scripts\hideusers.txt

foreach ($user in $users) {

    $samname = Get-ADUser -Filter {Name -eq $user} | Select-Object -Expand samaccountname
    
    Write-Host "Setting: $samname"
    set-aduser $samname -add @{msExchHideFromAddressLists='TRUE'}
    
   

}