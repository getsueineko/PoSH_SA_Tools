$searchOU = "OU=Subunit,OU=staff,DC=contoso,DC=com"
$lastdays = 365
$date = (get-date).AddDays(-$lastdays)
Get-ADComputer -SearchBase $searchOU -Filter {LastLogonTimeStamp -lt $date} -properties *| Select-Object name, LastLogonDate, DistinguishedName | ConvertTo-Html | Out-File "$home\Documents\Zombies.html"