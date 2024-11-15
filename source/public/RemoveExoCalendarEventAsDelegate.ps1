Function Remove-ExoCalendarEventAsDelegate {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High', DefaultParameterSetName = "ByInputObject")]
    param(
        [Parameter(Mandatory, ValueFromPipeline, ParameterSetName = "ByInputObject")]
        [PSCustomObject]  # Optionally specify a custom type if available
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

        # Define the helper function here in the begin block
        function Remove-CalendarEvent {
            param (
                [string] $UserId,
                [string] $EventIdentifier
            )
            if ($PSCmdlet.ShouldProcess($UserId, "Remove calendar event [$EventIdentifier]")) {
                try {
                    Remove-MgUserEvent -UserId $UserId -EventId $EventIdentifier -ErrorAction Stop
                    Write-Verbose "Successfully removed calendar event [$EventIdentifier] for user [$UserId]."
                }
                catch {
                    Write-Error "Failed to remove event [$EventIdentifier] for user [$UserId]: $_"
                }
            }
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

                # Extract properties and call helper function
                Remove-CalendarEvent -UserId $item.MailboxId -EventIdentifier $item.EventId
            }
        }
        elseif ($PSCmdlet.ParameterSetName -eq 'ById') {
            # Directly remove event using provided parameters
            Remove-CalendarEvent -UserId $MailboxId -EventIdentifier $EventId
        }
    }
    end {}
}
