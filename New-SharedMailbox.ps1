try {
    Connect-ExchangeOnline -ShowBanner:$false -CommandName New-Mailbox,Set-Mailbox,Set-MailboxRegionalConfiguration -Device -ErrorAction Stop
    Connect-MgGraph -Scopes "User.EnableDisableAccount.All" -NoWelcome -UseDeviceCode -ErrorAction Stop

    $name = Read-Host "Shared Mailbox name (eg: Sales Department)"
    $alias = Read-Host "Shared Mailbox alias (eg: Sales)"

    $smb = New-Mailbox -Shared -Name $name -DisplayName $name -Alias $alias -ErrorAction Stop
    Set-Mailbox -Identity $smb.ExternalDirectoryObjectId -MessageCopyForSentAsEnabled $true
    Write-Output "$name ($alias) shared mailbox created"

    Update-MgUser -UserId $smb.ExternalDirectoryObjectId -AccountEnabled:$false
    Write-Output "Blocked sign-in on the shared mailbox"

    $setToFrench = Read-Host "Set mailbox language to French? (Y/N)"
    if ($setToFrench.ToUpper() -eq "Y") {
        Write-Output "Setting mailbox to French..."
        Set-MailboxRegionalConfiguration -Identity $smb.ExternalDirectoryObjectId -Language fr-CA -LocalizeDefaultFolderName
    }
    else {
        Write-Output "Skipping French setup..."
    }
}
catch {
    Write-Output "Error:"
    Write-Output $_
}
finally {
    Write-Output "Disconnecting from EXO and Graph...."
    Disconnect-ExchangeOnline -Confirm:$false
    Disconnect-MgGraph
}
