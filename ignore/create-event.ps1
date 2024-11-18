$params = @{
    subject               = "Is this recurring?"
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
            startDate = "2024-05-01"
            endDate   = "2024-11-09"
        }
    }
    location              = @{
        displayName = "Harry's Bar"
    }
    attendees             = @(
        @{
            emailAddress = @{
                address = "dummy_user@poshlab.xyz"
                name    = "Dummy User"
            }
            type         = "required"
        }
        @{
            emailAddress = @{
                address = "mailer365@poshlab.xyz"
                name    = "Mailer365"
            }
            type         = "required"
        }
    )
    allowNewTimeProposals = $true
}

New-MgUserEvent -UserId mailer365@poshlab1.onmicrosoft.com -BodyParameter $params