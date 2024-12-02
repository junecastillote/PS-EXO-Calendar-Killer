# Function to remove calendar event
Function Remove-ExoCalendarEventAsDelegate {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = "High" , DefaultParameterSetName = "ByInputObject")]
    param(
        [Parameter(Mandatory, ValueFromPipeline, ParameterSetName = "ByInputObject")]
        [ValidateNotNullOrEmpty()]
        [System.Management.Automation.PSTypeName('PSEXOCalendarEvent')]
        $InputObject,

        [Parameter(Mandatory, ParameterSetName = "ById")]
        [string]
        $MailboxId,

        [Parameter(Mandatory, ParameterSetName = "ById")]
        [string]
        $EventId
    )
    begin {
        Write-Verbose "ParameterSet: $($PSCmdlet.ParameterSetName)"

        # Get the MS Graph logged in UPN value.
        $delegate_user = (Get-MgContext).Account

        # Initialize an arraylist object to hold mailboxid and eventid values.
        $event_to_remove = [System.Collections.ArrayList]@()

        Function AddToUserEventCollection {
            param (
                $MailboxId,
                $EventId
            )

            $null = $event_to_remove.Add(
                [pscustomobject](
                    [ordered]@{
                        MailboxId = $MailboxId
                        EventId   = $EventId
                    }
                )
            )
        }
    }

    process {
        # Process items based on the parameter set
        if ($PSCmdlet.ParameterSetName -eq 'ByInputObject') {
            foreach ($item in $InputObject) {
                # Ensure input object type matches
                if ($item.PSTypeNames -notcontains 'PSEXOCalendarEvent') {
                    throw "Input object is not of type PSEXOCalendarEvent. Received type: $($item.PSTypeNames -join ', ')"
                }
                AddToUserEventCollection $item.MailboxId $item.EventId
            }
        }
        elseif ($PSCmdlet.ParameterSetName -eq 'ById') {
            AddToUserEventCollection $MailboxId $EventId
        }
    }
    end {

        for ($i = 0; $i -lt $event_to_remove.Count; $i++) {
            $item = $event_to_remove[$i]
            if ($PSCmdlet.ShouldProcess($($item.MailboxId), "Remove calendar event [$($item.EventId)]")) {
                try {
                    if (!$current_mailbox_id) {
                        $current_mailbox_id = $item.MailboxId
                        if (!(AddCalendarPermission -MailboxId $item.MailboxId -DelegateId $delegate_user)) {
                            Continue
                        }
                    }

                    if ($current_mailbox_id -ne $item.MailboxId) {
                        RemoveCalendarPermission -MailboxId $current_mailbox_id -DelegateId $delegate_user
                        $current_mailbox_id = $item.MailboxId
                        Write-Verbose "Current Mailbox = $(($current_mailbox_id).ToUpper())"
                        if (!(AddCalendarPermission -MailboxId $item.MailboxId -DelegateId $delegate_user)) {
                            Continue
                        }
                    }

                    # Remove the calendar event
                    try {
                        Remove-MgUserEvent -UserId $item.MailboxId -EventId $item.EventId -ErrorAction Stop
                        Write-Verbose "[$($item.MailboxId)]: Successfully removed calendar event [$($item.EventId)]."
                    }
                    catch {
                        Write-Error "[$($item.MailboxId)]: Failed to remove the calendar event."
                        Write-Error "$($_.Exception.Message)"
                    }

                    # Remove calendar folder permission if it's the last item in the array.
                    if ($i -eq ($event_to_remove.Count - 1)) {
                        RemoveCalendarPermission -MailboxId $current_mailbox_id -DelegateId $delegate_user
                    }
                }
                catch {
                    Write-Error "[$($item.MailboxId)]: Failed to remove event [$($item.EventId)]."
                    Write-Error "$($_.Exception.Message)"
                }
            }
        }
    }
}