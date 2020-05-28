<#
.NOTES
      Author: Anatoly Evdokimov
      Date: August 11 2019
#>

# Очистить экран
Clear-Host
# Очистка переменной $Error
$Error.Clear()

#Import-Module BitsTransfer

$backupDir = '\\backupserver\d$\Archive\'
$pstBackup = '\\mailserver\pst$\'
$ExchServer = 'http://exchange.contoso.com/powershell'
$profileFolder = '\\contoso.com\root\usersdata\Subunit\Profiles\'
$userFolder = '\\contoso.com\root\usersdata\Subunit\UserFolders\'
$trigger = 0 

# --=== ФУНКЦИИ ===-- #

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

# Функция уведомления
function Add-Notify ([String]$aTitle, [String]$aTipText) {
    Add-Type –AssemblyName System.Speech
    $SpeechSynthesizer = New-Object –TypeName System.Speech.Synthesis.SpeechSynthesizer
    $SpeechSynthesizer.SelectVoice("Microsoft Irina Desktop")
    $SpeechSynthesizer.Speak($aTitle)

    #[console]::beep(659, 500)
    #[console]::beep(659, 500)
    #[console]::beep(659, 500)
    #[console]::beep(698, 350)
    #[console]::beep(523, 150)
    #[console]::beep(415, 500)
    #[console]::beep(349, 350)
    #[console]::beep(523, 150)
    #[console]::beep(440, 1000)

    Add-Type -AssemblyName  System.Windows.Forms
    $balloonActionResult = New-Object System.Windows.Forms.NotifyIcon
    $balloonActionResult.Icon = [System.Drawing.SystemIcons]::Information
    $balloonActionResult.BalloonTipTitle = $aTitle
    $balloonActionResult.BalloonTipIcon = "Info"
    $balloonActionResult.BalloonTipText = $aTipText 
    $balloonActionResult.Visible = $True
    $balloonActionResult.ShowBalloonTip(5000)
}

# Функция проверки существования и копирования папки
function Copy-ExistFolder ([String]$checkDir, [String]$deskordoc) {
    if (Test-Path -Path "$checkDir\$deskordoc") {
        Copy-Item "$checkDir\$deskordoc" -Destination "$backupDir$SamAccountName" -Recurse -Force -PassThru -Exclude "`$RECYCLE.BIN"
    }
    else {
        Write-Host "$checkDir\$deskordoc не найдена" -ForegroundColor DarkYellow
    }
}

# Функция бэкапа папок
function Backup-Folder ([String]$aFolderName) {
    $sizeFolder = "{0:N2} MB" -f ((Get-ChildItem –Force "$userFolder$SamAccountName\$aFolderName" –Recurse -ErrorAction SilentlyContinue | Measure-Object Length -s).sum / 1Mb)
    Write-Host "`nРазмер папки $userFolder$SamAccountName\$aFolderName - $sizeFolder" -ForegroundColor DarkYellow
    $approveBackupFolder = Approve-Action -varquestion "Делаем бэкап папки `«$aFolderName`»" -approve $trigger
    if ($approveBackupFolder -eq 1) {
        Copy-ExistFolder -checkDir "$userFolder$SamAccountName" -deskordoc "$aFolderName"
    }
}

# Функция проверки существования и удаления папки
function Remove-ExistFolder ([String]$checkDir) {
    if (Test-Path -Path $checkDir) {
        Write-Host "Удалось найти $checkDir. `nБудет удалено." -ForegroundColor DarkYellow
        Remove-Item –path $checkDir –Recurse -Force -Verbose
        Write-Host "Удалено. " -ForegroundColor DarkYellow
    }
    else {
        Write-Host "Не удалось найти $checkDir" -ForegroundColor DarkYellow
    }
}

# --=== ФУНКЦИИ ===-- #

$SamAccountName = read-host "Введите имя входа удаляемого пользователя(SAM)"
#Делаем простейшую проверку на существование такого юзверя
While (@(Get-ADUser $SamAccountName).count -ne 1) {
    Write-Host "`nЭто фиаско, братан! Нет такого пользователя, попробуй еще раз`n" -ForegroundColor DarkYellow
    $SamAccountName = read-host "Введите имя входа удаляемого пользователя(SAM)"
}

$a = Get-ADUser $SamAccountName | Get-ADObject -Properties lastLogon
$logonTime = [DateTime]::FromFileTime($a.lastLogon)
Write-Host "`nДанный пользователь логинился последний раз $logonTime`n" -ForegroundColor DarkYellow
Write-Host -nonewline "Хотите продолжить? [y/n] "
$response = read-host
if ( $response -ne "y" ) { exit }

$distinguishedName = Get-ADUser $SamAccountName -Properties DistinguishedName | ForEach-Object { $_.DistinguishedName }
$targetOU = $distinguishedName.Substring($distinguishedName.IndexOf('OU='))

#Выделяем фамилию пользователя
$InitialsUser = @((Get-ADUser $SamAccountName -Properties Name | ForEach-Object { $_.Name }).Split(" "))
$InitialsUser = $InitialsUser[0]

$checkProfileV2 = "$profileFolder$SamAccountName.V2"
$checkProfileV6 = "$profileFolder$SamAccountName.V6"

Remove-ExistFolder -checkDir $checkProfileV2
Remove-ExistFolder -checkDir $checkProfileV6

# Бэкап почты в PST
$approveMovePST = Approve-Action -varquestion "Делаем бэкап почты" -approve $trigger
if ($approveMovePST -eq 1) {
    do {
        $cred = Get-Credential "$env:USERDOMAIN\$env:USERNAME"
        $currentDomain = "LDAP://" + ([ADSI]"").distinguishedName
        $domain = New-Object System.DirectoryServices.DirectoryEntry($currentDomain, $cred.username, $cred.GetNetworkCredential().password)
    } until($null -ne $domain.name)

    $Session = New-PSSession -Authentication Kerberos -Credential $cred -ConnectionUri $ExchServer -ConfigurationName Microsoft.Exchange -SessionOption (New-PSSessionOption -SkipRevocationCheck)
    Import-PSSession $Session
    New-MailboxExportRequest -Mailbox $SamAccountName -FilePath "$pstBackup$SamAccountName.pst"
    $i = 1
    $iMax = Get-Mailbox -Identity $SamAccountName -resultsize unlimited | Get-MailboxStatistics
    $iMax = $iMax.TotalItemSize.Value
    $iMax = $iMax -replace "bytes", ""
    $iMax = $iMax -replace ",", ""
    $regex = [regex]"\((.*)\s\)"
    $iMax = [regex]::match($iMax, $regex).Groups[1]
    $iMax = [double]$iMax.Value
    $iMax = [math]::Round($iMax / [math]::Pow(1024, 2), 0) 
    do {
        Write-Progress -Activity "Перенос архива... Пожалуйста, не закрывайте программу до звукового сигнала и уведомления" -PercentComplete ($i / $iMax * 100) 
        $i = [math]::Round((Get-ChildItem –Force "$pstBackup$SamAccountName.pst" –Recurse -ErrorAction SilentlyContinue | Measure-Object Length -s).sum / 1Mb, 0)
        if ($i -gt $iMax) { $i = $iMax }
        Start-Sleep -Seconds 1
    } while (Get-MailboxExportRequest $_ | Where-Object { $_.Mailbox -match $InitialsUser } | Where-Object { $_.Status -match 'Queued|InProgress' })
    Remove-PSSession $Session
    
    Add-Notify -aTitle "Перемещение завершено" -aTipText "`nАрхив был экспортирован в `n$pstBackup$SamAccountName.pst"
    Start-Sleep -Seconds 10

    New-Item -ItemType directory -Path "$backupDir$SamAccountName"
    #Copy-Item "$pstBackup$SamAccountName.pst" -Destination "$backupDir$SamAccountName" -Force -Recurse -PassThru
    #Start-BitsTransfer –source "$pstBackup$SamAccountName.pst" -Destination "$backupDir$SamAccountName"
    robocopy $pstBackup "$backupDir$SamAccountName" "$SamAccountName.pst" /NP /ETA /R:1000 /W:30
}

if (-not (Test-Path -Path "$backupDir$SamAccountName")) {
    New-Item -ItemType directory -Path "$backupDir$SamAccountName"
}

# Бэкап папок Рабочий стол и Мои документы
Backup-Folder -aFolderName "My Documents"
Backup-Folder -aFolderName "Desktop"

# Окончательное удаление с рабочих ресурсов
$approveRemove = Approve-Action -varquestion "Удаляем все папки и учетку пользователя из AD" -approve $trigger
if ($approveRemove -eq 1) {
    # Снимаем защиту с контейнера
    Get-ADOrganizationalUnit –Identity $targetOU -Properties ProtectedFromAccidentalDeletion | 
    Set-ADOrganizationalUnit -ProtectedFromAccidentalDeletion $False 
    # Удаляем учетку
    # Используется конструкция с Remove-ADObject, чтобы решить проблему с удалением учеток содержащих другие объекты
    Get-ADUser -Identity $SamAccountName | Remove-ADObject -Recursive -Confirm:$false
    # Восстанавливаем защиту контейнера
    Get-ADOrganizationalUnit –Identity $targetOU -Properties ProtectedFromAccidentalDeletion | 
    Set-ADOrganizationalUnit -ProtectedFromAccidentalDeletion $True
       
    Remove-ExistFolder -checkDir "$userFolder$SamAccountName"
    Remove-ExistFolder -checkDir "$pstBackup$SamAccountName.pst"
}

# Чистим учетные данные
$cred = $null