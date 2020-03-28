<#
.NOTES
      Author: Anatoly Evdokimov
      Date: August 11 2019
#>

Clear-Host

$ErrorActionPreference = 'SilentlyContinue'

$UsersFolders = '\\contoso.com\root\usersdata\Subunit\UserFolders\'
$disOU = 'OU=Dismissed, OU=staff,DC=contoso,DC=com'
$ExchServer = 'http://exchange.contoso.com/powershell'
$smtp = "mail.contoso.com"
$from = "HelpDesk <helpdesk@contoso.com>"
$trigger = 0 

# --=== ФУНКЦИИ ===-- #

# Функция проверки на существование такого юзверя
function Test-AliveUser ([String]$aUsername, [String]$aInputext) {
    While (@(Get-ADUser $aUsername).count -ne 1) {
        Write-Host "`nЭто фиаско, братан! Нет такого пользователя, попробуй еще раз`n" -ForegroundColor DarkYellow
        $aUsername = read-host $aInputext
    }
    return $aUsername
}

# Функция выбора для последующего действия
function Approve-Action ([String]$varquestion, [String]$approve) {

    $question = read-host "$varquestion ? [y/n]"

    if ($question -eq "y") {
        $trigger = 1
        return $trigger
    } 
    ElseIf ($question -eq "n") {
        Write-Host "Ок"
        return $trigger
    } 
    Else {
        Write-Host "Ваш ответ был не [y/n], по умолчанию был выбран ответ n"
        return $trigger
    }
}

# Функция формирования и отправки письма
function Send-NotifyMessage () {    
    $to = $Leader + "@contoso.com"
    $subject = "Почтовый ящик уволенного сотрудника ($InitialsUser)"
    $body = "<BODY style=""font-size: 11pt; font-family: Arial""><P>"
    $body += "$salutation $InitialsLeaderFul!<br>"
    $body += " <br>"    
    $body += "В соответствии с регламентом по работе с ящиками электронной почты уволенных сотрудников,<br>"
    $body += "к Вашей учетной записи присоединен почтовый ящик уволенного сотрудника $InitialsUser сроком на 3 месяца.<br>"
    $body += "По всем вопросам работы с этим ящиком Вы можете обращаться в IT службу.<br>"
    $body += $addPermText
    $body += "</P></BODY>"    
    Send-MailMessage -From $from -To $to -Cc $from -Subject $subject -Body $body -BodyAsHtml -Encoding ([System.Text.Encoding]::UTF8) -SmtpServer $smtp
}

# --=== ФУНКЦИИ ===-- #

$SamAccountName = read-host "Введите имя входа удаляемого пользователя(SAM)"
$SamAccountName = Test-AliveUser -aUsername $SamAccountName -aInputext "Введите имя входа удаляемого пользователя(SAM)"
$Leader = read-host "Введите имя входа руководителя удаляемого пользователя(SAM)"
$Leader = Test-AliveUser -aUsername $Leader -aInputext "Введите имя входа руководителя удаляемого пользователя(SAM)"

$distinguishedName = Get-ADUser $SamAccountName -Properties DistinguishedName | ForEach-Object { $_.DistinguishedName }
$targetOU = $distinguishedName.Substring($distinguishedName.IndexOf('OU='))

# Выделяем ФИО пользователя
$InitialsUser = @((Get-ADUser $SamAccountName -Properties Name | ForEach-Object { $_.Name }).Split(" "))
$InitialsUser = $InitialsUser[0] + " " + ($InitialsUser[1])[0] + "." + " " + ($InitialsUser[2])[0] + "."

# Выделяем ИО руководителя
$InitialsLeader = @((Get-ADUser $Leader -Properties Name | ForEach-Object { $_.Name }).Split(" "))
$InitialsLeaderFull = $InitialsLeader[1] + " " + $InitialsLeader[2]

# Определяем пол руководителя
if ($InitialsLeader[2] -like "*овна" -or $InitialsLeader[2] -like "*евна" -or $InitialsLeader[2] -like "*ична") {
    $salutation = "Уважаемая"
}
else { 
    $salutation = "Уважаемый"
}


Disable-ADAccount -Identity $SamAccountName
# Снимаем защиту с контейнера
Get-ADOrganizationalUnit –Identity $targetOU -Properties ProtectedFromAccidentalDeletion | 
Set-ADOrganizationalUnit -ProtectedFromAccidentalDeletion $False 

# Перемещаем нашего героя
Get-ADUser $SamAccountName | Move-ADObject -TargetPath $disOU
 
# Восстанавливаем защиту контейнера
Get-ADOrganizationalUnit –Identity $targetOU -Properties ProtectedFromAccidentalDeletion | 
Set-ADOrganizationalUnit -ProtectedFromAccidentalDeletion $True

# Выдаем права руководителю на папку пользователя
$approvePermission = Approve-Action -varquestion "Предоставить руководителю доступ к папке пользователя" -approve $trigger
if ($approvePermission -eq 1) {
    cacls $UsersFolders\$SamAccountName /E /G ($Leader + ":F")
    $addPermText = "Доступ к рабочим документам предоставлен: <a href=$UsersFolders\$SamAccountName>$UsersFolders\$SamAccountName</a>"
}

# Делегируем права на ящик и мапим его руководителю
$approveMapping = Approve-Action -varquestion "Делегировать права на почтовый ящик" -approve $trigger
if ($approveMapping -eq 1) {
    do {
        $cred = Get-Credential "$env:USERDOMAIN\$env:USERNAME"
        $currentDomain = "LDAP://" + ([ADSI]"").distinguishedName
        $domain = New-Object System.DirectoryServices.DirectoryEntry($currentDomain, $cred.username, $cred.GetNetworkCredential().password)
    } until($null -ne $domain.name)

    $Session = New-PSSession -Authentication Kerberos -Credential $cred -ConnectionUri $ExchServer -ConfigurationName Microsoft.Exchange -SessionOption (New-PSSessionOption -SkipRevocationCheck)
    Import-PSSession $Session
    Add-mailboxpermission –Identity $SamAccountName –User $Leader –accessrights Fullaccess, readpermission –inheritancetype All –Automapping:$True
    Set-MailContact $SamAccountName -HiddenFromAddressListsEnabled $true
    Remove-PSSession $Session
}

# Отправляем письмо руководителю
$approveNotify = Approve-Action -varquestion "Отправить письмо руководителю" - approve $trigger

if ($approveNotify -eq 1) {
    Send-NotifyMessage
}

# Делаем себе в календаре напоминание
$approveAppointment = Approve-Action -varquestion "Сделать напоминание в личном календаре" - approve $trigger
if ($approveAppointment -eq 1) {
    $job = Start-Job -scriptblock {
        param ($InitialsUser)
        $ol = New-Object -ComObject Outlook.Application
        $meeting = $ol.CreateItem('olAppointmentItem')
        $meeting.Subject = 'Удаление пользователя'
        $meeting.Body = $InitialsUser
        $meeting.Location = 'Кабинет 932'
        $meeting.ReminderSet = $true
        $meeting.Importance = 1
        $meeting.MeetingStatus = [Microsoft.Office.Interop.Outlook.OlMeetingStatus]::olMeeting
        #$meeting.Recipients.Add('heldesk@vectorinfo.ru')
        $meeting.ReminderMinutesBeforeStart = 10080
        $meeting.ReminderMinutesBeforeStart = 15
        $meeting.Start = [datetime]::Today.Adddays(90)
        $meeting.Duration = 30
        $meeting.Send()
    } -Args $InitialsUser -credential $cred
}

# Чистим учетные данные
$cred = $null