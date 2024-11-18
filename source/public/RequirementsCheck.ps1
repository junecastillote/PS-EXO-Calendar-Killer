Function RequirementsCheck {
    [CmdletBinding()]
    param ()

    ## Test if the Microsoft Graph API is connected.
    if (!($context = Get-MgContext)) {
        "[FAIL] - Connect to Microsoft Graph First using the Connect-MgGraph cmdlet." | Write-Verbose
        return $false
    }
    "[PASS] - Microsoft Graph API Connected." | Write-Verbose

    ## Test if the Microsoft Graph API authentication type is "Delegated".
    if ($context.AuthType -ne 'Delegated') {
        "[FAIL] - This script must be run using Microsoft Graph API delegated authentication." | Write-Verbose
        return $false
    }
    "[PASS] - Microsoft Graph API Connection Delegated Access." | Write-Verbose

    ## Test if the required permissions are present in the Microsoft Graph API connection.
    $scopes_required = @('Calendars.ReadWrite', 'Calendars.ReadWrite.Shared')
    $scopes = @($context.Scopes | Where-Object { $scopes_required -contains $_ })
    if ($scopes.Count -lt $scopes_required.Count) {
        "[FAIL] - This script requires the following scopes: [$($scopes_required -join ',')]." | Write-Verbose
        return $false
    }
    "[PASS] - Microsoft Graph API Scopes [$($scopes_required -join ',')]." | Write-Verbose

    ## Test if the Microsoft Graph API logged in user has a mailbox
    $delegate_user = "$($context.Account)"
    try {
        $null = Get-MgUserDefaultCalendar -UserId $delegate_user -ErrorAction Stop
    }
    catch {
        if ($_.Exception.Message -like "*MailboxNotEnabledForRESTAPI*") {
            "[FAIL] - The Microsoft Graph API logged in user ($($delegate_user)) doesn't have a mailbox." | Write-Verbose

        }
        else {
            $_.Exception.Message | Write-Verbose
        }
        return $false
    }
    "[PASS] - The Microsoft Graph API logged in user ($($delegate_user)) has a valid mailbox." | Write-Verbose

    ## Test if EXO PowerShell is connected
    if ($null = Get-Module ExchangeOnlineManagement) {
        if (!($exoConn = (Get-ConnectionInformation))) {
            "[FAIL] - Connect to Exchange Online PowerShell first using the Connect-ExchangeOnline cmdlet." | Write-Verbose
            return $false
        }
        else {
            "[PASS] - Exchange Online PowerShell is connected." | Write-Verbose
        }
    }
    else {
        "[FAIL] - Connect to Exchange Online PowerShell first using the Connect-ExchangeOnline cmdlet." | Write-Verbose
        return $false
    }

    ## Test if the Exchange login account has the 'Mail Recipients' management role
    $exo_admin = $exoConn.UserPrincipalName
    try {
        $null = Get-Command Get-ManagementRoleAssignment -ErrorAction Stop
        if (!(Get-ManagementRoleAssignment -RoleAssignee $exo_admin -Delegating $false -Role "Mail Recipients")) {
            "[FAIL] - The Exchange Online PowerShell logged in user ($($exo_admin)) doesn't have the minimum required management role assignment [Mail Recipients]." | Write-Verbose
            return $false
        }
    }
    catch {
        "[FAIL] - The Exchange Online PowerShell logged in user ($($exo_admin)) doesn't have the minimum required management role assignment [Mail Recipients]." | Write-Verbose
        return $false
    }
    "[PASS] - The Exchange Online PowerShell logged in user ($($exo_admin)) have the minimum required management role assignment [Mail Recipients]." | Write-Verbose

    $env:DELEGATEISALLGOOD = $true
    $env:DELEGATEACCOUNT = $delegate_user
    $env:EXCHANGEADMINACCOUNT = $exo_admin
    return $true
}