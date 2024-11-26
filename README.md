# PS EXO Calendar Killer

PowerShell Module to find and delete orphaned calendar events

- [Required Modules](#required-modules)
- [Required Accounts](#required-accounts)
- [Preparing your shell](#preparing-your-shell)
- [Get-ExoCalendarEventAsDelegate Function](#get-exocalendareventasdelegate-function)
  - [Syntax - Get-ExoCalendarEventAsDelegate](#syntax---get-exocalendareventasdelegate)
  - [Parameters - Get-ExoCalendarEventAsDelegate](#parameters---get-exocalendareventasdelegate)
  - [Example - Search the target mailbox calendar events by organizer name](#example---search-the-target-mailbox-calendar-events-by-organizer-name)
  - [Example - Get all calendar events from the target mailbox](#example---get-all-calendar-events-from-the-target-mailbox)
  - [Example - Get all calendar events within a time range](#example---get-all-calendar-events-within-a-time-range)
- [Remove-ExoCalendarEventAsDelegate Function](#remove-exocalendareventasdelegate-function)
  - [Syntax - Remove-ExoCalendarEventAsDelegate](#syntax---remove-exocalendareventasdelegate)
  - [Parameters - Remove-ExoCalendarEventAsDelegate](#parameters---remove-exocalendareventasdelegate)
  - [Example - Delete found calendar events](#example---delete-found-calendar-events)
  - [Example - Delete calendar events by EventId](#example---delete-calendar-events-by-eventid)

## Required Modules

- [PowerShell EXO Calendar Killer Module](https://github.com/junecastillote/PS-EXO-Calendar-Killer/archive/refs/heads/main.zip)
- [Microsft Graph API PowerShell Module](https://learn.microsoft.com/en-us/powershell/microsoftgraph/installation)
- [Exchange Online Management (EXOV3) PowerShell Module](https://learn.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2?view=exchange-ps#install-and-maintain-the-exchange-online-powershell-module)

## Required Accounts

- User account with a valid mailbox (licensed).
  - This account will log in to Microsft Graph API PowerShell.
  - Admin role not required.
- User account with a minimum of 'Mail Recipients' management role in Exchange Online.
  - This account will log in to the Exchange Online PowerShell.

> **Note**: You can use the same or two separate user accounts for Exchange Online and Graph API PowerShell, as long as they comply with the requirements.
> Some organizations do not allow admin accounts to have a mailbox, instead, admins are given a normal user account with a mailbox.

## Preparing your shell

Before using the `ps-exo-calendar-killer` module, you must be connected to Graph and EXO sessions.

```powerShell
# Log in as a mailbox user (admin not required) to Microsoft Graph
Connect-MgGraph -TenantId 'TENANT.onmicrosoft.com' -Scopes Calendars.ReadWrite.Shared, Calendars.ReadWrite

# Log in as Exchange Admin to Exchange Online PowerShell
Connect-ExchangeOnline -Organization 'TENANT.onmicrosoft.com' -UserPrincipalName user@contoso.com
```

## Get-ExoCalendarEventAsDelegate Function

This function performs the following actions:

1. Add permission to the user calendar.
2. Search the calendar for all or specific events.
3. Remove calendar permission.

### Syntax - Get-ExoCalendarEventAsDelegate

```PowerShell
Get-ExoCalendarEventAsDelegate -MailboxId <string> -EventId <string> [-SkipRequirementsCheck] [<CommonParameters>]

Get-ExoCalendarEventAsDelegate -MailboxId <string> [-Subject <string>] [-StartDate <datetime>] [-EndDate <datetime>] [-OrganizerName <string>] [-ExactSubjectMatch <bool>] [-SkipRequirementsCheck] [<CommonParameters>]

Get-ExoCalendarEventAsDelegate -MailboxId <string> -FilterString <string> [-SkipRequirementsCheck] [<CommonParameters>]
```

### Parameters - Get-ExoCalendarEventAsDelegate

`-MailboxId`

The user principal name of the mailbox to search.

ie. `user@contoso.com`

|                            |        |
| -------------------------- | ------ |
| Type                       | String |
| Required                   | True   |
| Accept pipeline input      | False  |
| Accept wildcard characters | False  |
| Default value              | None   |

`-StartDate`

Specify the oldest start date of the event to include in the search.

|                            |          |
| -------------------------- | -------- |
| Type                       | DateTime |
| Required                   | False    |
| Accept pipeline input      | False    |
| Accept wildcard characters | False    |
| Default value              | None     |

`-EndDate`

Specify the newest end date of the event to include in the search.

|                            |          |
| -------------------------- | -------- |
| Type                       | DateTime |
| Required                   | False    |
| Accept pipeline input      | False    |
| Accept wildcard characters | False    |
| Default value              | None     |

`-OrganizerName`

The organizer name of the event to search.

ie. "John Doe"

|                            |        |
| -------------------------- | ------ |
| Type                       | String |
| Required                   | False  |
| Accept pipeline input      | False  |
| Accept wildcard characters | False  |
| Default value              | None   |

`-Subject`

The subject of the calendar event to search.

|                            |        |
| -------------------------- | ------ |
| Type                       | String |
| Required                   | False  |
| Accept pipeline input      | False  |
| Accept wildcard characters | False  |
| Default value              | None   |

`-ExactSubjectMatch`

Specify whether to perform an exact match of the subject.

For example:

`-Subject "Hello"` will search an exact match of the subject "`Hello`" by default. 

To perform an approximate subject match search, specify the `-Subject "Hello" -ExactSubjectMatch:$false`

|                            |         |
| -------------------------- | ------- |
| Type                       | Boolean |
| Required                   | False   |
| Accept pipeline input      | False   |
| Accept wildcard characters | False   |
| Default value              | True    |

### Example - Search the target mailbox calendar events by organizer name

> Note: This search works only if the organizer still exists in the organization / directory. If not, the result will return this error - `[ErrorItemNotFound] | the specified object was not found in the store.`

This example searches the calender for events organized by the user `John Doe` and stores the result to the `$calendar_events` variable.

```PowerShell
Get-ExoCalendarEventAsDelegate -MailboxId user@contoso.com -OrganizerName "John Doe" -OutVariable calendar_events
```

### Example - Get all calendar events from the target mailbox

> Note: Do not run this option needlessly because it will retrieve all calendar events without filter.

```PowerShell
Get-ExoCalendarEventAsDelegate -MailboxId user@contoso.com -OutVariable calendar_events
```

### Example - Get all calendar events within a time range

This example searches for calendar events whose start and end dates fall between 30 and 20 days ago.

> Note: The `StartDate` and `EndDate` refence pertains to the Start and End of the event, not the time they were created. The `StartDate` and `EndDate` parameters can be used together or individually.

```PowerShell
Get-ExoCalendarEventAsDelegate -MailboxId user@contoso.com -OutVariable calendar_events -StartDate (Get-Date).AddDays(-30) -EndDate (Get-Date).AddDays(-20)
```

## Remove-ExoCalendarEventAsDelegate Function

This function deletes the calendar event based on the spcified EventId.

### Syntax - Remove-ExoCalendarEventAsDelegate

```PowerShell
Remove-ExoCalendarEventAsDelegate -InputObject <PSEXOCalendarEvent> [-WhatIf] [-Confirm] [<CommonParameters>]

Remove-ExoCalendarEventAsDelegate -MailboxId <string> -EventId <string> [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Parameters - Remove-ExoCalendarEventAsDelegate

`-InputObject`

The `PSEXOCalendarEvent` output of the `Get-ExoCalendarEventAsDelegate` command.

|                            |                      |
| -------------------------- | -------------------- |
| Type                       | <PSEXOCalendarEvent> |
| Required                   | True                 |
| Accept pipeline input      | True                 |
| Accept wildcard characters | False                |
| Default value              | None                 |

`-MailboxId`

The user principal name of the mailbox containing the calendar event to delete.

ie. `user@contoso.com`

|                            |        |
| -------------------------- | ------ |
| Type                       | String |
| Required                   | True   |
| Accept pipeline input      | False  |
| Accept wildcard characters | False  |
| Default value              | None   |

`-EventId`

The specific event id of the calendar item to delete.

ie. `AQMkAGNhOTJlYzAyLTk2YjctNDllMC1iMmU2LTE3...`

|                            |        |
| -------------------------- | ------ |
| Type                       | String |
| Required                   | True   |
| Accept pipeline input      | False  |
| Accept wildcard characters | False  |
| Default value              | None   |

`-FilterString`

Specify the custom filter string.

```powerShell
$filter = "subject eq 'Is this recurring?' and organizer/emailaddress/name eq 'June Castillote'"
Get-ExoCalendarEventAsDelegate -MailboxId user@contoso.com -OutVariable calendar_event -Verbose -FilterString $filter
```

|                            |        |
| -------------------------- | ------ |
| Type                       | String |
| Required                   | False  |
| Accept pipeline input      | False  |
| Accept wildcard characters | False  |
| Default value              | None   |

### Example - Delete found calendar events

```PowerShell
# Passing the items through the pipeline input
$calendar_events | Remove-ExoCalendarEventAsDelegate

# Using the -InputObject Parameter
Remove-ExoCalendarEventAsDelegate -InputObject $calendar_events
```

### Example - Delete calendar events by EventId

Suppose you need to remove a specific event from the target calendar, you can do so by providing the `EventId` and `MailboxId` values.

```PowerShell
Remove-ExoCalendarEventAsDelegate -MailboxId user@contoso.com -EventId "AQMkAGNhOTJlYzAyLTk2YjctNDllM..."
```
