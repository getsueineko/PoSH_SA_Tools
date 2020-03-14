<#
.NOTES
      Author: Anatoly Evdokimov
      Date: September 17 2019
#>

# Очистить экран
Clear-Host
# Очистка переменной $Error
$Error.Clear()

$SamAccountName = read-host "Введите логин пользователя, который необходимо поменять"

[string]$SqlServer = "YOURSQLSERVERNAME";
[string]$SqlCatalog = "IntranetDB";
[string]$SqlLogin = "operator";
[string]$SqlPassw = "p@$$w0rd"
[string]$SQLQuery = $("SELECT [ItemID],[FirstName],[MiddleName],[LastName],[LatinName] FROM [IntranetDB].[dbo].[tm_Workers] WHERE [LatinName] = '$SamAccountName'")

$Connection = New-Object System.Data.SqlClient.SqlConnection
$Connection.ConnectionString = "Server=$SqlServer; Database=$SqlCatalog; User ID=$SqlLogin; Password=$SqlPassw;"
$Connection.Open()
$Command = New-Object System.Data.SQLClient.SQLCommand
$Command.Connection = $Connection
$Command.CommandText = $SQLQuery
$Reader = $Command.ExecuteReader()
Write-Host "`nПо вашему запросу найдены следующие пользователи: " 
while ($Reader.Read()) {
    $ItemID = $Reader.GetValue(0)
    $FirstName = $Reader.GetValue(1)
    $MiddleName = $Reader.GetValue(2)
    $LastName = $Reader.GetValue(3)
    $LatinName = $Reader.GetValue(4)
    Write-Host "ItemID: $ItemID, Имя: $FirstName, Отчество: $MiddleName, Фамилия: $LastName, Логин: $LatinName"
}
$Connection.Close()

$ItemIDAlter = read-host "`nВведите ItemID исправляемого пользователя"
$LatinNameAlter = read-host "Введите правильный вариант логина"
[string]$SQLQueryAlter = $("UPDATE [IntranetDB].[dbo].[tm_Workers] SET LatinName = '$LatinNameAlter' WHERE [ItemID] = '$ItemIDAlter'")

$Connection.Open()
$Command = New-Object System.Data.SQLClient.SQLCommand
$Command.Connection = $Connection
$Command.CommandText = $SQLQueryAlter
$Command.ExecuteNonQuery()
$Connection.Close()