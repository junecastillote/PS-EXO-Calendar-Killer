Function Get-ExoCalendarEventAsDelegate {
    [cmdletbinding()]
    param (
        [parameter(Mandatory)]
        [string[]]
        $MailboxId,

        [parameter(Mandatory, ParameterSetName = 'ByEventId')]
        [string]
        $EventId,

        [parameter(ParameterSetName = 'FilterOption')]
        [string]
        $Subject,

        [parameter(ParameterSetName = 'FilterOption')]
        [datetime]
        $StartDate,

        [parameter(ParameterSetName = 'FilterOption')]
        [datetime]
        $EndDate,

        [parameter(ParameterSetName = 'FilterOption')]
        [string]
        $OrganizerName,

        [parameter(Mandatory, ParameterSetName = 'FilterString')]
        [string]
        $FilterString,

        [parameter(ParameterSetName = 'FilterOption')]
        [bool]
        $ExactSubjectMatch = $true,

        [Parameter()]
        [switch]
        $SkipDelegateRequirementsCheck
    )

    ## Function to split strings.
    ## Example: "seriesMaster" to "Series Master"
    function Format-String ($str) {
        # Convert the first character to uppercase
        $str = $str.Substring(0, 1).ToUpper() + $str.Substring(1)

        # Insert a space before each uppercase character
        $str = [regex]::Replace($str, '([A-Z])', ' ${1}')

        # Remove leading space
        $str = $str.TrimStart()

        return $str
    }

    # $startTime = Get-Date

    if ($PSVersionTable.PSEdition -eq 'Core') {
        $PSStyle.Progress.View = 'Classic'
    }

    if (!$SkipDelegateRequirementsCheck) {
        if (!(Confirm-DelegateRequirements)) {
            "Prerequisites check failed." | Write-Error
            return $null
        }
    }

    $delegate_user = (Get-MgContext).Account

    if ($FilterString) {
        $search_filter = $FilterString
    }

    if ($PSCmdlet.ParameterSetName -eq 'FilterOption') {
        $filter_option_array = [System.Collections.Generic.List[string]]::new()

        if ($Subject) {
            ## If exact match (subject)
            if ($ExactSubjectMatch -eq $true) {
                $filter_option_array.Add(
                    $("Subject eq '$($Subject)'")
                )
            }
            ## If approximate match (subject)
            else {
                $filter_option_array.Add(
                    $("contains(subject,'$($Subject)')")
                )
            }
        }

        if ($OrganizerName) {
            $filter_option_array.Add(
                "organizer/emailaddress/name eq '$($OrganizerName)'"
            )
        }

        if ($StartDate) {
            $filter_option_array.Add(
                "start/datetime ge '$((Get-Date $StartDate).ToString('yyyy-MM-ddTHH:mm:ss'))'"
            )
        }

        if ($EndDate) {
            $filter_option_array.Add(
                "end/datetime lt '$((Get-Date $EndDate).ToString('yyyy-MM-ddTHH:mm:ss'))'"
            )
        }

        if ($filter_option_array.Count -gt 1) {
            $filter_option_array = $filter_option_array -join " and "
        }

        $search_filter = $filter_option_array
    }

    for ($i = 0; $i -lt ($MailboxId.Count) ; $i++) {
        $current_mailbox = $MailboxId[$i]
        ## Test if the target user has a mailbox.
        try {
            $null = Get-Mailbox -Identity $current_mailbox -ErrorAction Stop
        }
        catch {
            if ($_.Exception.Message -like "Ex6F9304*") {
                "The mailbox for [$($current_mailbox)] is not found." | Write-Error
            }
            else {
                $_.Exception.Message | Write-Error
            }
            return $null
        }

        ## Give your delegate user Editor permission to the target mailbox calendar folder.
        "[$($current_mailbox)]: Adding [$($delegate_user)] permission to calendar." | Write-Verbose
        $null = Add-MailboxFolderPermission -Identity "$($current_mailbox):\Calendar" -User $delegate_user -AccessRights Editor

        ## Define the filter. In this case, filter by event organizer name and subject.
        $searchSplat = @{
            # UserId = $current_mailbox
        }

        if ($PSCmdlet.ParameterSetName -eq 'ByEventId') {
            $searchSplat.Add('EventId', $EventId)
        }

        if ($search_filter) {
            $searchSplat.Add('Filter', $search_filter)
            $searchSplat.Add('All', $true)
            "Filter = $search_filter" | Write-Verbose
        }

        # $searchSplat | Out-Default
        ## Get all calendar events based on the given filter and store the result (if any) to the $cal_event variable.
        "[$($current_mailbox)]: Searching calendar." | Write-Verbose
        # $cal_event = @(Get-MgUserEvent -UserId $current_mailbox -All -Filter $filter)
        try {
            $cal_event = @(Get-MgUserEvent -UserId $current_mailbox @searchSplat -ErrorAction Stop)
        }
        catch {
            $_.Exception.Message | Write-Error
            # return $null
        }


        if ($cal_event.Count -lt 1) {
            "[$($current_mailbox)]: Found 0 calendar events." | Write-Verbose
            # return $null
        }

        if ($cal_event.Count -gt 0) {
            ## Preview the result
            "[$($current_mailbox)]: Found $($cal_event.Count) calendar events." | Write-Verbose
            # $result = $cal_event | ForEach-Object {
            $cal_event | ForEach-Object {
                $item = $_ | Select-Object `
                @{n = 'EventId'; e = { $_.Id } },
                @{n = 'MailboxId'; e = { $current_mailbox } },
                Subject, Categories,
                @{n = 'Type'; e = { $(Format-String $_.Type) } },
                @{n = 'OrganizerEmail'; e = { $_.Organizer.EmailAddress.Address } },
                @{n = 'OrganizerName'; e = { $_.Organizer.EmailAddress.Name } },
                @{n = 'CreatedDateTime'; e = { Get-Date $_.CreatedDateTime } },
                @{n = 'Start'; e = { (Get-Date $_.Start.DateTime) } },
                @{n = 'StartTimeZone'; e = { ($_.Start.TimeZone) } },
                @{n = 'OriginalStart'; e = { (Get-Date $_.OriginalStart) } },
                @{n = 'OriginalStartTimeZone'; e = { ($_.OriginalStartTimeZone) } },
                @{n = 'End'; e = { (Get-Date $_.End.DateTime) } },
                @{n = 'EndTimeZone'; e = { ($_.End.TimeZone) } },
                @{n = 'OriginalEnd'; e = { (Get-Date $_.OriginalEnd) } },
                @{n = 'OriginalEndTimeZone'; e = { ($_.OriginalEndTimeZone) } },
                @{n = 'LastModifiedDateTime'; e = { Get-Date $_.LastModifiedDateTime } },
                @{n = 'Attendees'; e = { ($_.Attendees | ConvertTo-Json -Depth 5) } },
                @{n = 'Recurrence'; e = { ($_.Recurrence.Pattern | ConvertTo-Json -Depth 5) } },
                @{n = 'Content'; e = { ($_.Body.Content) } },
                @{n = 'ContentType'; e = { ($_.Body.ContentType) } },
                @{n = 'Preview'; e = { ($_.BodyPreview) } },
                @{n = 'ResponseStatus'; e = { ($_.ResponseStatus | ConvertTo-Json -Depth 5) } },
                @{n = 'Locations'; e = { ($_.Locations | ConvertTo-Json -Depth 5) } },
                HasAttachments, ResponseRequested, ShowAs, TransactionId, Sensitivity, ReminderMinutesBeforeStart, IsReminderOn,
                IsOrganizer, IsOnlineMeeting, IsDraft, IsCancelled, IsAllDay, Importance, AllowNewTimeProposals, HideAttendees,
                OnlineMeetingProvider, OnlineMeetingUrl, WebLink

                # Add custom type
                $item.PSObject.TypeNames.Insert(0, 'PSEXOCalendarEvent')
                $item
            }
        }

        # Cleanup permission
        "[$($current_mailbox)]: Removing [$($delegate_user)] permission to calendar." | Write-Verbose
        $null = Remove-MailboxFolderPermission -Identity "$($current_mailbox):\Calendar" -User $delegate_user -Confirm:$false
    }
}