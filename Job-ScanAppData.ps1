$date = Get-Date -Format MM-dd-yy
$path = "\\desktop01\temp\scan\$env:USERNAME-$date$path.txt"
$apps = Get-childitem -Recurse "$Env:USERPROFILE\AppData\Roaming\*" -Include "*.exe" | select -ExpandProperty fullname

if (!(Test-Path $path)) {
Add-Content $apps -Path $path -Force

}

