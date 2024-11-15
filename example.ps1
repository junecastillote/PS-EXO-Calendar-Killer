## Log in as a mailbox user (admin not required) to Microsoft Graph
Connect-MgGraph -TenantId 'TENANT.onmicrosoft.com' -Scopes Calendars.ReadWrite.Shared, Calendars.ReadWrite

## Log in as Exchange Admin to Exchange Online PowerShell
Connect-ExchangeOnline -Organization 'TENANT.onmicrosoft.com' -UserPrincipalName user@contoso.com

## Import the module
Import-Module .\ps-exo-calendar-killer.psd1 -Force

## Example 1: Search the target mailbox for calendar events created by the user "John Doe" and store in $calendar_events variable.
Get-ExoCalendarEventAsDelegate -MailboxId user@contoso.com -OrganizerName "John Doe" -OutVariable calendar_events