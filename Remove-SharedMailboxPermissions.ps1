try {
    Connect-ExchangeOnline -ShowBanner:$false -CommandName Get-EXOMailbox,Get-EXOMailboxFolderStatistics,Remove-MailboxFolderPermission,Remove-MailboxPermission,Remove-RecipientPermission,Set-Mailbox -Device -ErrorAction Stop

    $smb = Read-Host "Shared mailbox (eg: shared@company.com)"
    $user = Read-Host "User (eg: mshaevitch@company.com)"

    Write-Output "Shared mailbox:"
    Get-EXOMailbox -Identity $smb -ErrorAction Stop
    Write-Output "User:"
    Get-EXOMailbox -Identity $user -ErrorAction Stop

    $continue = Read-Host "OK to continue? (Y/N)"

    if ($continue.ToUpper() -ne "Y") {
        Write-Output "Canceled by user. Exiting..."
        exit 0
    }

    Write-Output "Continuing..."

    $calendarFolder = Get-EXOMailboxFolderStatistics -Identity $smb -FolderScope Calendar | Where-Object { $_.FolderType -eq "Calendar" }
    $calendarUPN = $smb + ":\" + $calendarFolder.Name
    Remove-MailboxFolderPermission -Identity $calendarUPN -User $user -Confirm:$false -ErrorAction Continue

    Remove-MailboxPermission -Identity $smb -User $user -AccessRights FullAccess -InheritanceType All -Confirm:$false -ErrorAction Continue
    Remove-RecipientPermission -Identity $smb -Trustee $user -AccessRights SendAs -Confirm:$false -ErrorAction Continue
    Set-Mailbox -Identity $smb -GrantSendOnBehalfTo @{Remove=$user} -Confirm:$false -ErrorAction Continue

     Write-Output "Done! $user removed from $smb"
}
catch {
    Write-Error "An error occurred: $_"
}
finally {
    Write-Output "Disconnecting from EXO...."
    Disconnect-ExchangeOnline -Confirm:$false
}
