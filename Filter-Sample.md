# Custom Filter Examples

```powerShell
## Subject only (OK)
$filter = "subject eq 'Is this recurring?'"
$filter = "subject eq 'Just a sample event'"
## Subject + organizer name (OK)
$filter = "subject eq 'Is this recurring?' and organizer/emailaddress/name eq 'June Castillote'"
## Subject + start (or end) (NOT OK)
$filter = "subject eq 'Just a sample event' and start/datetime ge '2024-05-10T00:00:00'"
$filter = "contains(subject,'Is this recurring?') and start/datetime ge '2024-05-01T00:00:00'"
$filter = "subject eq 'Is this recurring?' and start/datetime ge '2024-05-01T00:00:00'"
## Organizer name + start (OK)
$filter = "organizer/emailaddress/name eq 'June Castillote' and start/datetime ge '2024-05-10T00:00:00'"
## Organizer name + start + end (OK)
$filter = "organizer/emailaddress/name eq 'June Castillote' and start/datetime ge '2024-05-01T00:00:00' and end/datetime le '2024-05-20T00:00:00'"
## Start + end (OK)
$filter = "start/datetime ge '2024-05-01T00:00:00' and end/datetime le '2024-05-20T00:00:00'"
```
