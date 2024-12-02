

$params = @{
    subject               = "Macy's Birthday"
    body                  = @{
        contentType = "HTML"
        content     = "Does noon time work for you?"
    }
    start                 = {
        dateTime = "05/01/2024 00:00:00"
        timeZone = "UTC"
    }
    end                   = {
        dateTime = "05/01/2024 00:30:00"
        timeZone = "UTC"
    }
    recurrence            = @{
        pattern = @{
            type       = "weekly"
            interval   = 1
            daysOfWeek = @(
                "Wednesday"
            )
        }
        range   = @{
            type      = "endDate"
            startDate = "2024-12-25"
            endDate   = "2024-12-26"
        }
    }
    location              = @{
        displayName = "Harry's Bar"
    }
    attendees             = @(
        @{
            emailAddress = @{
                address = "june@poshlab.xyz"
                name    = "June Castillote"
            }
            type         = "required"
        }
    )
    allowNewTimeProposals = $true
}

$null = New-MgUserEvent -UserId mailer365@poshlab.xyz -BodyParameter $params


# get-exoCalendarEventAsDelegate -MailboxId mailer365@poshlab.xyz -OutVariable calendar_events -Subject "Macy''s Birthday" -Verbose | remove-ExoCalendarEventAsDelegate -Verbose