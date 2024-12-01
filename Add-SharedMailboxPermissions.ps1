try {
    Connect-ExchangeOnline -ShowBanner:$false -CommandName Get-EXOMailbox,Add-MailboxPermission,Add-RecipientPermission,Add-MailboxFolderPermission,Get-EXOMailboxFolderStatistics -Device -ErrorAction Stop

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

    # Assign Read and Manage permissions
    Add-MailboxPermission -Identity $smb -User $user -AccessRights FullAccess -InheritanceType All

    # Assign Send As permissions
    Add-RecipientPermission -Identity $smb -AccessRights SendAs -Trustee $user -Confirm:$false

    # Assign permissions to view private items
    $calendarFolder = Get-EXOMailboxFolderStatistics -Identity $smb -FolderScope Calendar | Where-Object { $_.FolderType -eq "Calendar" }
    $calendarUPN = $smb + ":\" + $calendarFolder.Name
    Add-MailboxFolderPermission -Identity $calendarUPN -User $user -AccessRights Editor -SharingPermissionFlags Delegate,CanViewPrivateItems
    Write-Output "$calendarUPN calendar permissions assigned to $user"
}
catch {
    Write-Error "An error occurred: $_"
}
finally {
    Write-Output "Disconnecting from EXO...."
    Disconnect-ExchangeOnline -Confirm:$false
}
