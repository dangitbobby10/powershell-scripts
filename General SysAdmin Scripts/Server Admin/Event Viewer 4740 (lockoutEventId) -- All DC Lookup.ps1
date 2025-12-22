# Determines which Domain Controller has the recorded Lockout Event ID 4740 for a specified user.
# Note: You might need to adjust permissions and ensure remote event log access is enabled on the domain controllers.
#------------------------------------------------------------------------------------------------------------------------
# Define the username and domain to search
$userToSearch = Read-Host -Prompt 'Enter the account username'
$domain = "" #contoso.local

# Get the list of domain controllers in the domain
$domainControllers = Get-ADDomainController -Filter * -Server $domain

# Define the event ID for account lockouts
$lockoutEventId = 4740

# Loop through each domain controller and search for lockout events
foreach ($dc in $domainControllers) {
    # Query the security log for account lockout events related to the specified user
    $lockoutEvents = Get-WinEvent -ComputerName $dc.HostName -FilterHashtable @{
        LogName='Security'
        ID=$lockoutEventId
        StartTime=(Get-Date).AddDays(-1) # Adjust time range as needed
    } -ErrorAction SilentlyContinue | Where-Object {
        $_.Properties[0].Value -eq $userToSearch
    }

    # Display the results
    foreach ($event in $lockoutEvents) {
        $eventTime = $event.TimeCreated
        $lockedAccount = $event.Properties[0].Value
        $lockingComputer = $event.Properties[1].Value

        Write-Host "Lockout event for $lockedAccount occurred at $eventTime on DC $($dc.HostName). Lockout source: $lockingComputer"
    }
}