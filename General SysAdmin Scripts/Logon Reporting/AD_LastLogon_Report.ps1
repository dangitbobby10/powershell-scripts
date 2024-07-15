$session = New-PSSession -ComputerName "" #servername
Invoke-Command -Session $session -ScriptBlock {
    $logdate = get-date -f MM-dd-yyyy
    $csvfile = "c:\users\$env:username\desktop\Roster_AD_Queries_AD_Only-$logdate.csv"

    # Get a list of every DC
    $dcNames = Get-ADDomainController -Filter * |
      Select-Object -ExpandProperty Name |
      Sort-Object

    # Filter OU and by Active Users
    $searchBase = "DC=contoso,DC=com" #update to your domain
    $users = Get-ADUser -Filter 'enabled -eq $true' -Properties * -SearchBase $searchBase | Where-Object {$_.info -NE 'Migrated'}

    $output = foreach ($user in $users) {
        # Clear variables
        $latestLogonFT = $latestLogonServer = $latestLogon = $null
        $latestInteractiveLogonFT = $latestInteractiveLogon = $null  # Initialize variable for interactive logon time

        # Iterate every DC name
        foreach ($dcName in $dcNames) {
            # Query specific DC
            $userDetails = Get-ADUser -Server $dcName -Identity $user.SamAccountName -Properties lastLogon, msDS-LastSuccessfulInteractiveLogonTime |
                Select-Object -Property lastLogon, msDS-LastSuccessfulInteractiveLogonTime

            $lastLogonFT = $userDetails.lastLogon
            $interactiveLogonFT = $userDetails.'msDS-LastSuccessfulInteractiveLogonTime'

            # Remember most recent file time and DC name for lastLogon
            if ($lastLogonFT -and ($lastLogonFT -gt $latestLogonFT)) {
                $latestLogonFT = $lastLogonFT
                $latestLogonServer = $dcName
            }
            # Check for the most recent interactive logon time
            if ($interactiveLogonFT -and ($interactiveLogonFT -gt $latestInteractiveLogonFT)) {
                $latestInteractiveLogonFT = $interactiveLogonFT
            }
        }

        if ($latestLogonFT -and ($latestLogonFT -gt 0)) {
            # If user ever logged on, get DateTime from file time
            $latestLogon = [DateTime]::FromFileTime($latestLogonFT)
        } else {
            # User never logged on
            $latestLogon = $latestLogonServer = $null
        }

        if ($latestInteractiveLogonFT -and ($latestInteractiveLogonFT -gt 0)) {
            $latestInteractiveLogon = [DateTime]::FromFileTime($latestInteractiveLogonFT)
        }

        # Output User Data
        $user | Select-Object `
            @{Label = "Display Name"; Expression = {$_.displayname}},
            @{Name = "LatestLogon"; Expression = {$latestLogon}},
            @{Name = "LatestLogonServer"; Expression = {$latestLogonServer}},
            @{Name = "LatestInteractiveLogon"; Expression = {$latestInteractiveLogon}},  # Add interactive logon time
            @{Label = "User Access Control"; Expression = {$_.userAccountControl}},
            @{Label = "User Principal Name"; Expression = {$_.userPrincipalName}},
            @{Label = "Employee ID"; Expression = {$_.employeeID}},
            @{Label = "First Name"; Expression = {$_.GivenName}},
            @{Label = "Last Name"; Expression = {$_.Surname}},
            @{Label = "Manager"; Expression = {%{(Get-AdUser $_.Manager -server $dcName -Properties DisplayName).DisplayName}}},
            @{Label = "Organizational Unit"; Expression = {$_.distinguishedName}}
    }

    $output | Export-Csv -Path $csvfile -NoTypeInformation
}
