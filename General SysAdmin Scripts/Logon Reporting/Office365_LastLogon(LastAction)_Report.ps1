#Connect/Signin to Exchange Online
Connect-ExchangeOnline

$logdate = get-date -f MM-dd-yyyy
$csvfile = "c:\users\$env:username\desktop\Office365_LastLogon_Report-$logdate.csv"

# Start measuring time
$StartTime = Get-Date

# Get all mailboxes
$Mailboxes = Get-Mailbox -ResultSize Unlimited

# Export Displayname, UPN, PrimarySMTP, LastUserActionTime, TotalMailboxSize, User or Shared Mailbox, DirSync Status, Sign In Status
$Mailboxes | Select-Object DisplayName, UserPrincipalName, PrimarySmtpAddress, @{Name="LastUserActionTime";Expression={(Get-MailboxStatistics $_).LastUserActionTime}}, @{Name="LastUserAccessTime";Expression={(Get-MailboxStatistics $_).LastUserAccessTime}}, @{Name="LastLogonTime";Expression={(Get-MailboxStatistics $_).LastLogonTime}}, @{Name="TotalMailboxSize";Expression={(Get-MailboxStatistics $_).TotalItemSize}}, @{Name="User or Shared Mailbox";Expression={(Get-Mailbox $_).RecipientTypeDetails}}, @{Name="IsDirSynced";Expression={(Get-Mailbox $_).IsDirSynced}}, @{Name="Sign In Status";Expression={if ($_.ExchangeUserAccountControl -eq "none") { "Enabled" } else { "Blocked" }}} | Export-Csv -Path "$csvfile" -NoTypeInformation

# Stop measuring time
$EndTime = Get-Date

# Calculate and display the script runtime
$ElapsedTime = $EndTime - $StartTime
Write-Host "The script ran for $($ElapsedTime.TotalMinutes) minutes."