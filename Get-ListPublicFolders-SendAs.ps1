If (Test-Path MailPublicFolderUserPermission1.txt) { Remove-Item MailPublicFolderUserPermission1.txt }
If (Test-Path MailPublicFolderUserPermission2.txt) { Remove-Item MailPublicFolderUserPermission2.txt }
Write-Output $("Public folder path`tPublic folder DisplayName`tSend-As users") | out-file -Append MailPublicFolderUserPermission1.txt
Write-Output $("Public folder path`tPublic folder DisplayName`tSend-As users") | out-file -Append MailPublicFolderUserPermission2.txt

$PFList = Get-PublicFolder -Recurse | Where-Object { $_.MailEnabled -eq "True" };
foreach ($PF in $PFList) {
    $MPF = Get-MailPublicFolder -Identity $PF;
    $userEntry = @();
    $userList = '';
    $PFID = "\" + $PF.Name.Replace('"', '""').Replace('&', '`&');
    if ($PF.ParentPath -ne "\") { $PFID = $PF.ParentPath + $PFID };

    # Показать общие папки, которые видны в адресной книге
    if (-Not $MPF.HiddenFromAddressListsEnabled) {
        $permissions = Get-ADPermission -Identity $MPF.Identity
        foreach ($user in $permissions) {
            # Если нужно каждого пользователя в одной строке, то оставляем первое условие и комментируем второе.
            # Если нужен список всех пользователей в одной строке, то оставляем второе условие и комментируем первое.
            if ( $user.Extendedrights -like 'Send-As' -and $user.User -notlike 'S-1-5-21-*') { Write-Output ($PFID + "`t" + $MPF.DisplayName + "`t" + $user.User) | out-file -Append MailPublicFolderUserPermission1.txt };
            if ( $user.Extendedrights -like 'Send-As' -and $user.User -notlike 'S-1-5-21-*') { $userEntry += $user.User };
        };

        if ($userEntry.length -gt 0) {
            $userList = $($userEntry -join ','); 
            Write-Output $($PFID + "`t" + $MPF.DisplayName + "`t" + $userList)  | out-file -Append MailPublicFolderUserPermission2.txt;
        };
    }
}