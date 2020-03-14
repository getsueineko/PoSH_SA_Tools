<#
.SYNOPSIS
.DESCRIPTION
.EXAMPLE
.INPUTS
      None
.OUTPUTS
      None
.NOTES
      Author: Anatoly Evdokimov
      Date: April 11 2017
#>

$t = '[DllImport("user32.dll")] public static extern bool ShowWindow(int handle, int state);'
add-type -name win -member $t -namespace native
[native.win]::ShowWindow(([System.Diagnostics.Process]::GetCurrentProcess() | Get-Process).MainWindowHandle, 0)

[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Windows.Forms.Application]::EnableVisualStyles()

#Add-Type -AssemblyName System.Windows.Forms
#Add-Type -AssemblyName System.Drawing

#Создаем GUI и выполняем главный код
function FormAndScan {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Kill Zombies 2.0.7"
    $form.Size = New-Object System.Drawing.Size(300, 200)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = 'Fixed3D'
    $form.MaximizeBox = $false

    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = New-Object System.Drawing.Point(75, 125)
    $OKButton.Size = New-Object System.Drawing.Size(75, 23)
    $OKButton.Text = "OK"
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $OKButton
    $form.Controls.Add($OKButton)

    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = New-Object System.Drawing.Point(150, 125)
    $CancelButton.Size = New-Object System.Drawing.Size(75, 23)
    $CancelButton.Text = "Отмена"
    $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $CancelButton
    $form.Controls.Add($CancelButton)

    $labelCont = New-Object System.Windows.Forms.Label
    $labelCont.Location = New-Object System.Drawing.Point(10, 20)
    $labelCont.Size = New-Object System.Drawing.Size(280, 20)
    $labelCont.Text = "Введите имя контейнера для поиска:"
    $form.Controls.Add($labelCont) 

    $textBoxCont = New-Object System.Windows.Forms.TextBox
    $textBoxCont.Location = New-Object System.Drawing.Point(10, 40)
    $textBoxCont.Size = New-Object System.Drawing.Size(260, 20)
    $textBoxCont.name = "textBoxCont"
    $textBoxCont.add_MouseHover($ShowHelp)
    $form.Controls.Add($textBoxCont)

    $labelRange = New-Object System.Windows.Forms.Label
    $labelRange.Location = New-Object System.Drawing.Point(10, 70)
    $labelRange.Size = New-Object System.Drawing.Size(280, 20)
    $labelRange.Text = "Введите глубину поиска:"
    $form.Controls.Add($labelRange)

    $textBoxRange = New-Object System.Windows.Forms.TextBox
    $textBoxRange.Location = New-Object System.Drawing.Point(10, 90)
    $textBoxRange.Size = New-Object System.Drawing.Size(260, 20)
    $textBoxRange.name = "textBoxRange"
    $textBoxRange.add_MouseHover($ShowHelp)
    $form.Controls.Add($textBoxRange)

    $form.Topmost = $True

    $form.Add_Shown( {$textBoxCont.Select()})
    $result = $form.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $global:searchOU = $textBoxCont.Text
        if (!$searchOU) { 
            $searchOU = "OU=IT,OU=staff,DC=contoso,DC=com"
        }
        $global:lastdays = $textBoxRange.Text
        if (!$lastdays) {
            $lastdays = 365
        }
    }

    #здесь tooltip был

    if ($result -eq [System.Windows.Forms.DialogResult]::Cancel) {
        exit
    }

    #Ядро программы :D
    $global:date = (get-date).AddDays(-$lastdays)
    $global:mainscan = Get-ADComputer -SearchBase $searchOU -Filter {LastLogonTimeStamp -lt $date} -properties *| Select name, LastLogonDate, DistinguishedName | ConvertTo-Html | Out-File "$home\Documents\Zombies.html"

}

$tooltip1 = New-Object System.Windows.Forms.ToolTip
$ShowHelp = {
    Switch ($this.name) {
        "textBoxCont" {$tip = "По умолчанию: OU=IT,OU=staff,DC=contoso,DC=com"}
        "textBoxRange" {$tip = "По умолчанию: 365 (дней)"}
    }

    $tooltip1.SetToolTip($this, $tip)
} 

#Парсим ошибку как можем
try {
    FormAndScan
}
catch {
    $mainscan = $null 
    [System.Windows.Forms.MessageBox]::Show("Введены некорректные данные или нет такого OU", '', 'OK', 'Warning')
    FormAndScan
}

$msgBoxInput = [System.Windows.Forms.MessageBox]::Show("Результаты сканирования были сохранены в файл $home\Documents\Zombies.html `n`nХотите создать контейнер OU и перенести в него найденные компьютеры?", '', 'YesNo', 'Info')
switch ($msgBoxInput) {
    'Yes' {
        $dname = Get-ADDomain | % { $_.Forest }
        $dname = $dname.Split("{.}")
        $dnameroot = $dname[0]
        $dnamesuff = $dname[-1]
        $parentOU = "DC=$dnameroot,DC=$dnamesuff" 
        $navnpaaou = 'Zombies'
        $newOU = "OU=$navnpaaou,$parentOU"

        if (Get-ADOrganizationalUnit -Filter "distinguishedName -eq '$newOU'") {
            #Write-Host "Раздел AD $newOU уже существует."
            Get-ADComputer -SearchBase $searchOU -Filter {LastLogonTimeStamp -lt $date} -properties * | Move-ADObject -TargetPath $newOU
        }
        else {
            New-ADOrganizationalUnit -Name $navnpaaou -Path $parentOU
            #Write-Host "Новый раздел AD $newOU успешно создан."
            Get-ADComputer -SearchBase $searchOU -Filter {LastLogonTimeStamp -lt $date} -properties * | Move-ADObject -TargetPath $newOU
        }

        $balloonActionResult = New-Object System.Windows.Forms.NotifyIcon 
        $balloonActionResult.Icon = [System.Drawing.SystemIcons]::Information
        $balloonActionResult.BalloonTipTitle = "Перемещение завершено"
        $balloonActionResult.BalloonTipIcon = "Info"
        $balloonActionResult.BalloonTipText = "`nНайденные хосты были перемещены в $newOU"
        $balloonActionResult.Visible = $True
        $balloonActionResult.ShowBalloonTip(5000)
    }
    'No' {
        & $home\Documents\Zombies.html
        Start-Sleep -Seconds 1
        $balloonScanResult = New-Object System.Windows.Forms.NotifyIcon 
        $balloonScanResult.Icon = [System.Drawing.SystemIcons]::Information
        $balloonScanResult.BalloonTipTitle = "Сканирование завершено"
        $balloonScanResult.BalloonTipIcon = "Info"
        $balloonScanResult.BalloonTipText = "`nРезультаты сканирования были сохранены в файл $home\Documents\Zombies.html"
        $balloonScanResult.Visible = $True
        $balloonScanResult.ShowBalloonTip(5000)
    }
}