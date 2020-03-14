Clear-Host

#$ErrorActionPreference = 'SilentlyContinue'

$Password = 'P@$$w0rd'
$UsersFolders = '\\contoso.com\root\usersdata\Subunit\UsersFolders\'
$ProfilePath = '\\contoso.com\root\usersdata\Subunit\Profiles\'
$ExchServer = 'http://exchange.contoso.com/powershell'
$pikachu = 'http://intranet/helpdesktool/'

#Копируем с данного пользователя права
$TemplateUser = read-host "Введите имя входа шаблонного пользователя (SAM)"

$SamAccountName = read-host "Введите имя входа создаваемого пользователя(SAM)"
#Делаем простейшую проверку на существование такого юзверя
While (Get-ADUser -filter {SamAccountName -eq $SamAccountName}) {
    Write-Host "`nЭто фиаско, братан! Есть уже такой пользователь, попробуй еще раз`n" -ForegroundColor DarkYellow
    $SamAccountName = read-host "Введите имя входа пользователя (SAM)"
}

$ProfilePath = $ProfilePath + $SamAccountName
$UPN = $SamAccountName + '@contoso.com'

$DisplayName = read-host "Введите отображаемое имя (Display Name)"

[string]$SqlServer = "YOURSQLSERVERNAME";
[string]$SqlCatalog = "IntranetDB";
[string]$SqlLogin = "operator";
[string]$SqlPassw = "p@$$w0rd"
[string]$SQLQuery = $("SELECT [FirstName],[MiddleName],[LastName],[LatinName],[Birthday],[Email],[PhoneOffice],[PhoneExt],[RoomNumber],[StateListPost],[StateListDepartment],[Domen] FROM [IntranetDB].[dbo].[tm_Workers] WHERE [LatinName] = '$SamAccountName'")

$Connection = New-Object System.Data.SqlClient.SqlConnection
$Connection.ConnectionString = "Server=$SqlServer; Database=$SqlCatalog; User ID=$SqlLogin; Password=$SqlPassw;"
$Connection.Open()
$Command = New-Object System.Data.SQLClient.SQLCommand
$Command.Connection = $Connection
$Command.CommandText = $SQLQuery
$Reader = $Command.ExecuteReader()
while ($Reader.Read()) {
    $FirstName = $Reader.GetValue(0)
    $MiddleName = $Reader.GetValue(1)
    $LastName = $Reader.GetValue(2)
    # $LatinName = $Reader.GetValue(3)
    $Birthday = $Reader.GetValue(4)
    $DomainSmtpAddress = $Reader.GetValue(5)
    $PhoneOffice = $Reader.GetValue(6)
    # $PhoneExt = $Reader.GetValue(7)
    $RoomNumber = $Reader.GetValue(8)
    $Post = $Reader.GetValue(9)
    # $Dptmt = $Reader.GetValue(10)
    $Mail = $Reader.GetValue(11)
}
$Connection.Close()

$Post = $Post.substring(0,1).toupper()+$Post.substring(1)
$PrimarySmtpAddress = $SamAccountName + "@" + $Mail
$Dptmt = Get-ADUser $TemplateUser -Properties Department | ForEach-Object { $_.Department }

function Test-DataDB ([String]$varF, [String]$nameF) {

    $choiceF = read-host "Из базы Интранета $nameF был определен как $varF, это верно? [y/n]"

    if ($choiceF -eq "y") {
        return $varF
    } 
    ElseIf ($choiceF -eq "n") {
        $varF = read-host "Введите правильный $nameF пользователя"
        return $varF
    } 
    Else {
        Write-Host "`Ваш ответ был не [y/n], поэтому в качестве $nameF остался $varF"
        return $varF
    }
}

$PrimarySmtpAddress = Test-DataDB -varF $PrimarySmtpAddress -nameF "PrimarySmtpAddress"
$Dptmt = Test-DataDB -varF $Dptmt -nameF "отдел"

[string]$b = [datetime]$Birthday
$b = $b.Remove(10)
$arr3 = $b -split "`/"
[string]$Info = $arr3[1] + "." + $arr3[0] + "." + $arr3[2]
if (!$Info) {$Info = '01.01.1970'}

$Org = read-host "Введите название компании"
$Descript = $Post + ", " + $Org

$FullName = $LastName + " " + $FirstName + " " + $MiddleName

$arr2 = $DisplayName -split ' '
$DisplayNameMail = $arr2[1] + " " + $arr2[0]

#Выделяем OU
$regex_dn = '^CN=(?<cn>.+?)(?<!\\),(?<ou>(?:(?:OU|CN).+?(?<!\\),)+(?<dc>DC.+?))$'
$dn = Get-ADUser $TemplateUser | ForEach-Object { $_.DistinguishedName }
$dn -match $regex_dn
$TargetOU = $Matches['ou']

#Создаем пользователя
New-ADUser -Name $FullName -GivenName $FirstName -Surname $LastName -SamAccountName $SamAccountName -UserPrincipalName $UPN -DisplayName $DisplayName -Path $TargetOU -AccountPassword (ConvertTo-SecureString $Password -AsPlainText -force) -Enabled $true -ChangePasswordAtLogon $true -ProfilePath $ProfilePath -Description $Descript -Department $Dptmt -Title $Post -Company $Org -OfficePhone $PhoneOffice -Office $RoomNumber -EmailAddress $DomainSmtpAddress
$NewUser = $SamAccountName
Set-ADUser $NewUser -replace @{info = $Info}

Start-Sleep -s 5

#Копируем членство, прости господи, из шаблонного пользователя
$CopyFromUser = Get-ADUser $TemplateUser -prop MemberOf
$CopyToUser = Get-ADUser $NewUser -prop MemberOf
$CopyFromUser.MemberOf | Where-Object {$CopyToUser.MemberOf -notcontains $_} |  Add-ADGroupMember -Members $CopyToUser -PassThru

#Создаем папку пользователя и даем ему доступ на нее
mkdir $UsersFolders\$NewUser
#cacls $UsersFolders\$NewUser /E /G ($NewUser + ":F")
icacls $UsersFolders\$NewUser /grant ($NewUser + ":(OI)(CI)(F)")
$check_permission = icacls $UsersFolders\$NewUser

do {
    $cred = Get-Credential "$env:USERDOMAIN\$env:USERNAME"
    $currentDomain = "LDAP://" + ([ADSI]"").distinguishedName
    $domain = New-Object System.DirectoryServices.DirectoryEntry($currentDomain, $cred.username, $cred.GetNetworkCredential().password)
} until($null -ne $domain.name)

$Session = New-PSSession -Authentication Kerberos -Credential $cred -ConnectionUri $ExchServer -ConfigurationName Microsoft.Exchange -SessionOption (New-PSSessionOption -SkipRevocationCheck)
Import-PSSession $Session
Enable-Mailbox -Identity $UPN -Alias $NewUser
Set-Mailbox $UPN -EmailAddressPolicyEnabled $false -PrimarySmtpAddress $PrimarySmtpAddress -SimpleDisplayName $DisplayNameMail
Remove-PSSession $Session

Get-ADUser $SamAccountName -ErrorAction SilentlyContinue 
If (!($?)) {
    Write-Host "Не удается найти объект с удостоверением: $SamAccountName в DC.`nЧто-то пошло не так, проверьте вводимые данные.`n" -ForegroundColor Red
}
else {
    Write-Host "На папку пользователя выставлены следующие разрешения: $check_permission[0] `n" -ForegroundColor Green

    Write-Host "Пользователь создан. Проверьте, что все данные верные.`n" -ForegroundColor Green
    
    $choice = read-host "Если все верно, то может вы хотите помочь Пикачу (~‾▿‾)~ [y/n]"

    if ($choice -eq "y") {
        $WebResponse = Invoke-WebRequest $pikachu -Credential $cred
        $filter = "($LastName)\D+\d+.\d+.\d+\D+(\d+)"
        $catch = [regex]::Matches($WebResponse.Content, $filter).Groups[2].Value
        $link = $WebResponse.Links | Where-Object {$_.href -Match $catch} | Select-Object -Property href | ForEach-Object { $_.href } # Если на странице много одинаковых ссылок, то переменная будет типа массив и дергаем по индексу
        Start-Process ($pikachu + $link)
    } 
    Else {
        Write-Host "Ну что ж, придется открыть IE и потрудиться вместе с Пикачу (╥_╥)"
    }
}

$cred = $null

