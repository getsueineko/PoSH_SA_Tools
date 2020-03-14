Clear-Host

#$ErrorActionPreference = 'SilentlyContinue'

$Password = 'P@$$w0rd'
$UsersFolders = '\\contoso.com\root\usersdata\Subunit\UsersFolders\'
$ProfilePath = '\\contoso.com\root\usersdata\Subunit\Profiles\'
$ExchServer = 'http://exchange.contoso.com/powershell'

#Копируем с данного пользователя права
$TemplateUser = read-host "Введите имя входа шаблонного пользователя (SAM)"

$SamAccountName = read-host "Введите имя входа создаваемого пользователя (SAM)"
#Делаем простейшую проверку на существование такого юзверя
While (Get-ADUser -filter {SamAccountName -eq $SamAccountName}) {
    Write-Host "`nЭто фиаско, братан! Есть уже такой пользователь, попробуй еще раз`n" -ForegroundColor DarkYellow
    $SamAccountName = read-host "Введите имя входа пользователя (SAM)"
}

$ProfilePath = $ProfilePath + $SamAccountName
$UPN = $SamAccountName + '@contoso.com'
$FullName = read-host "Введите ФИО пользователя"
$DisplayName = read-host "Введите отображаемое имя (Display Name)"

$Info = read-host "Введите дату рождения пользователя (по умолчанию 01.01.1970)"
#Дефолтим переменную, если она пустая
if (!$Info) {$Info = '01.01.1970'}

$PrimarySmtpAddress = read-host "Введите PrimarySmtpAddress пользователя"
$Dptmt = read-host "Введите название отдела"
$Post = read-host "Введите название должности"
$Org = read-host "Введите название компании"
$Descript = $Post + ", " + $Org

$arr = $FullName -split ' '
[string]$FirstName = $arr[0]
[string]$LastName = $arr[1]

$arr2 = $DisplayName -split ' '
$DisplayNameMail = $arr2[1] + " " + $arr2[0]

#Выделяем OU
$regex_dn = '^CN=(?<cn>.+?)(?<!\\),(?<ou>(?:(?:OU|CN).+?(?<!\\),)+(?<dc>DC.+?))$'
$dn = Get-ADUser $TemplateUser | ForEach-Object { $_.DistinguishedName }
$dn -match $regex_dn
$TargetOU = $Matches['ou']

#Создаем пользователя
New-ADUser -Name $FullName -GivenName $FirstName -Surname $LastName -SamAccountName $SamAccountName -UserPrincipalName $UPN -DisplayName $DisplayName -Path $TargetOU -AccountPassword (ConvertTo-SecureString $Password -AsPlainText -force) -Enabled $true -ChangePasswordAtLogon $true -ProfilePath $ProfilePath -Description $Descript -Department $Dptmt -Title $Post -Company $Org
$NewUser = $SamAccountName
Set-ADUser $NewUser -replace @{info = $Info}

#Копируем членство, прости господи, из шаблонного пользователя
$CopyFromUser = Get-ADUser $TemplateUser -prop MemberOf
$CopyToUser = Get-ADUser $NewUser -prop MemberOf
$CopyFromUser.MemberOf | Where-Object {$CopyToUser.MemberOf -notcontains $_} |  Add-ADGroupMember -Members $CopyToUser -PassThru

#Создаем папку пользователя и даем ему доступ на нее
mkdir $UsersFolders\$NewUser
#cacls $UsersFolders\$NewUser /E /G ($NewUser + ":F")
icacls $UsersFolders\$NewUser /grant ($NewUser + ":(OI)(CI)(F)")

$Session = New-PSSession -Authentication Kerberos -Credential (Get-Credential) -ConnectionUri $ExchServer -ConfigurationName Microsoft.Exchange -SessionOption (New-PSSessionOption -SkipRevocationCheck)
Import-PSSession $Session
Enable-Mailbox -Identity $UPN -Alias $NewUser
Set-Mailbox $UPN -EmailAddressPolicyEnabled $false -PrimarySmtpAddress $PrimarySmtpAddress -SimpleDisplayName $DisplayNameMail
Remove-PSSession $Session

Get-ADUser $SamAccountName -ErrorAction SilentlyContinue 
If (!($?)) {
    Write-Host "Не удается найти объект с удостоверением: $SamAccountName в DC.`nЧто-то пошло не так, проверьте вводимые данные.`n" -ForegroundColor Red
}
else {
    Write-Host "Пользователь создан. Проверьте, что все данные верные." -ForegroundColor Green
}