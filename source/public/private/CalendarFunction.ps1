# Function to add calendar folder permission
Function AddCalendarPermission {
    [CmdletBinding()]
    param(
        [string]$MailboxId,
        [string]$DelegateId
    )

    ## Give yourself delegate user Editor permission to the target mailbox calendar folder.
    "[$($MailboxId)]: Adding [$($DelegateId)] permission to calendar." | Write-Verbose
    try {
        $null = Add-MailboxFolderPermission -Identity "$($MailboxId):\Calendar" -User $DelegateId -AccessRights Editor -ErrorAction Stop
        return $true
    }
    catch {
        if ($_.Exception.Message -like "*Microsoft.Exchange.Management.StoreTasks.UserAlreadyExistsInPermissionEntryException*") {
            Write-Verbose $_.Exception.Data.Values.Message
            return $true
        }
        else {
            Write-Error "Failed to add calendar permission. Cannot search this mailbox for events."
            Write-Error $_.Exception.Message
            return $false
        }
    }
}

# Function to remove calendar folder permission
Function RemoveCalendarPermission {
    [CmdletBinding()]
    param(
        [string]$MailboxId,
        [string]$DelegateId
    )

    # Cleanup permission
    "[$($MailboxId)]: Removing [$($DelegateId)] permission to calendar." | Write-Verbose
    $null = Remove-MailboxFolderPermission -Identity "$($MailboxId):\Calendar" -User $DelegateId -Confirm:$false -WarningAction SilentlyContinue
}