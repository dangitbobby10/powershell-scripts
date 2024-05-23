# The idea being you need to monitor a server/computer's logon events without 3rd party software. Complete this with a Scheduled Task in Windows.

# Define the path to the CSV file
$logPath = ""

# Get the current date and time
$currentDateTime = Get-Date

# Define the time span for new events (e.g., the last 1 hour)
$timeSpan = New-TimeSpan -Hours 1

# Calculate the start time for the query
$startTime = $currentDateTime.AddHours(-1)

# Query the event log for successful (4624), unsuccessful (4625) logon events, and logout events (4634) in the last hour
$events = Get-WinEvent -FilterHashtable @{LogName='Security'; Id=4624,4625,4634; StartTime=$startTime} | ForEach-Object {
    $event = [xml]$_.ToXml()
    $username = $event.Event.EventData.Data | Where-Object {$_.Name -eq 'TargetUserName'} | Select-Object -ExpandProperty '#text'
    if ($username -notmatch '^DWM-' -and $username -notmatch '^UMFD-' -and $username -notin @('SYSTEM', 'CCBQVWEHR01$')) {
        [PSCustomObject]@{
            TimeGenerated = $_.TimeCreated
            EntryType     = $_.Id
            InstanceId    = $_.Id
            UserName      = $username
            Message       = $_.Message
        }
    }
}

# Filter out null entries and export the events to a CSV file
$events | Where-Object { $_ -ne $null } | Export-Csv -Path $logPath -Append -NoTypeInformation