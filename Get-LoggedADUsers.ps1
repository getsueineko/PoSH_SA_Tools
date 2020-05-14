Clear-Host

$ErrorActionPreference = 'SilentlyContinue'

$i = 0
$computers = (Get-ADComputer -Filter * -SearchBase "OU=staff,DC=contoso,DC=com").Name
$result = foreach ($computer in $computers) {
    $descript = (Get-ADComputer -Identity $computer -Properties Description).Description
    If (Test-Connection -BufferSize 32 -Count 1 -ComputerName $computer -Quiet) {
        Invoke-Command -ComputerName $computer -ArgumentList $descript -ScriptBlock {
            param($descript)
            $props = @{PCName = $env:COMPUTERNAME }
            $props.Add('Description', $descript)
            $explorerprocesses = @(Get-WmiObject -Query "Select * FROM Win32_Process WHERE Name='explorer.exe'" -ErrorAction SilentlyContinue)
            If ($explorerprocesses.Count -eq 0) {
                $props.Add('User', $false)  
                $props.Add('State', 'No explorer process found / Nobody interactively logged on') 
            }
            Else {
                ForEach ($i in $explorerprocesses) {
                    $Username = $i.GetOwner().User
                    $Domain = $i.GetOwner().Domain
                    $state = "Logged on since: $($i.ConvertToDateTime($i.CreationDate))"
                    $props.Add('User', "$Domain\$Username")
                    $props.Add('State', $state) 
                }        
            }
            New-Object -Type PSObject -Prop $props              
        }
        Write-Progress -Activity "Retrieving data from $computer which is $i out of $($computers.Count)" -percentComplete  ($i++*100/$computers.Count) -status Running
    }    
}
Write-Host "A total of $i from $($computers.Count) computers were processed" -ForegroundColor Yellow 
$result | Select-Object PCName, Description, User, State | Out-GridView
#$result | Select-Object PCName, Description, User, State | Where-Object User -Like "*Smirnov | Format-Table