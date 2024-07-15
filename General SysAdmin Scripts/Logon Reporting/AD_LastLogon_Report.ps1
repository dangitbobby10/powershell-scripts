$logdate = get-date -f MM-dd-yyyy
$csvfile = "c:\temp\AD_LastLogon_Report-$logdate.csv" #update your path

# Get a list of every DC
$dcNames = Get-ADDomainController -Filter * |
  Select-Object -ExpandProperty Name |
  Sort-Object

# Filter OU
$searchBase = "DC=contoso,DC=com" #update with your domain

###########################################################################################################
# Search All Users
$users = Get-ADUser -Filter * -Properties * -SearchBase $searchBase | Where-Object {$_.info -NE 'Migrated'}

# Specify a single user
#$users = Get-ADUser -Identity "test.user" -Properties *
###########################################################################################################

# Hashtable used for splatting for Get-ADUser in loop
$params = @{
  "Properties" = @("lastLogon", "msDS-LastSuccessfulInteractiveLogonTime")
}

$output = foreach ( $user in $users ) {
  # Set LDAPFilter to find specific user
  $params.LDAPFilter = "(sAMAccountName=$($user.SamAccountName))"

  # Clear variables
  $latestLogonFT = $latestLogonServer = $latestLogon = $null
  $latestInteractiveLogonFT = $latestInteractiveLogon = $null

  # Iterate every DC name
  foreach ( $dcName in $dcNames ) {
    # Query specific DC
    $params.Server = $dcName

    # Get lastLogon and msDS-LastSuccessfulInteractiveLogonTime attributes
    $userDetails = Get-ADUser @params |
      Select-Object -Property lastLogon, msDS-LastSuccessfulInteractiveLogonTime

    $lastLogonFT = $userDetails.lastLogon
    $interactiveLogonFT = $userDetails.'msDS-LastSuccessfulInteractiveLogonTime'

    # Remember most recent file time and DC name for lastLogon
    if ( $lastLogonFT -and ($lastLogonFT -gt $latestLogonFT) ) {
      $latestLogonFT = $lastLogonFT
      $latestLogonServer = $dcName
    }

    # Check for the most recent interactive logon time
    if ($interactiveLogonFT -and ($interactiveLogonFT -gt $latestInteractiveLogonFT)) {
      $latestInteractiveLogonFT = $interactiveLogonFT
    }
  }

  if ( $latestLogonFT -and ($latestLogonFT -gt 0) ) {
    # If user ever logged on, get DateTime from file time
    $latestLogon = [DateTime]::FromFileTime($latestLogonFT)
  }
  else {
    # User never logged on
    $latestLogon = $latestLogonServer = $null
  }

  if ($latestInteractiveLogonFT -and ($latestInteractiveLogonFT -gt 0)) {
    $latestInteractiveLogon = [DateTime]::FromFileTime($latestInteractiveLogonFT)
  }

  # Get lastLogonTimestamp attribute
  $lastLogonTimestamp = $user.lastLogonTimestamp

  # Convert lastLogonTimestamp to DateTime if it's not null
  if ($lastLogonTimestamp -ne $null) {
    $lastLogonTimestamp = [DateTime]::FromFileTime($lastLogonTimestamp)
  }

  # Output User Data
  $user | Select-Object `
    @{Label = "Display Name";Expression = {$_.name}},
    @{Label = "Password Last Set";Expression = {$_.PasswordLastSet}},
    @{Label = "Password Expires";Expression = {if ($_.PasswordNeverExpires -eq $false) {$_.PasswordLastSet.AddDays(90)} else {"Never"}}}, # Assuming password expires in 90 days
    @{Label = "LatestLogon";Expression = {$latestLogon}},
    @{Label = "LatestLogonServer"; Expression = {$latestLogonServer}},
    @{Label = "LastLogonTimestamp"; Expression = {$lastLogonTimestamp}},
    @{Label = "LatestInteractiveLogon"; Expression = {$latestInteractiveLogon}},  # Add interactive logon time
    @{Label = "User Principal Name";Expression = {$_.userPrincipalName}},
    @{Label = "Pre-Windows 2000 Username";Expression = {$_.sAMAccountName}},
    @{Label = "Employee ID";Expression = {$_.employeeID}},
    @{Label = "First Name";Expression = {$_.GivenName}},
    @{Label = "Last Name";Expression = {$_.Surname}},
    @{Label = "Manager";Expression = {%{(Get-AdUser $_.Manager -Properties DisplayName).DisplayName}}},
    @{Label = "Distinguished Name";Expression = {$_.DistinguishedName}},
    @{Label = "Organizational Unit";Expression = {($_.DistinguishedName -split '(?<=,)',2)[1] -replace 'DC=','' -replace ',','.'}},
    @{Label = "Account Status";Expression = {if ($_.Enabled -eq $true) {"Enabled"} else {"Disabled"}}}
}

$output | Export-Csv -Path $csvfile -NoTypeInformation
