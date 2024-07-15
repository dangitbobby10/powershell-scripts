#   ___   __  __ _            _____  __  ____       _             _ _ _     ____                       _   
#  / _ \ / _|/ _(_) ___ ___  |___ / / /_| ___|     / \  _   _  __| (_) |_  |  _ \ ___ _ __   ___  _ __| |_ 
# | | | | |_| |_| |/ __/ _ \   |_ \| '_ \___ \    / _ \| | | |/ _` | | __| | |_) / _ \ '_ \ / _ \| '__| __|
# | |_| |  _|  _| | (_|  __/  ___) | (_) |__) |  / ___ \ |_| | (_| | | |_  |  _ <  __/ |_) | (_) | |  | |_ 
#  \___/|_| |_| |_|\___\___| |____/ \___/____/  /_/   \_\__,_|\__,_|_|\__| |_| \_\___| .__/ \___/|_|   \__|
#                                                                                    |_|                   

<# --------------------------------------------------------------------------------------------------------------------------------------------------------
Written by dangitbobby10
The organization I made this script for is using a 3rd party application for MFA so checking if MFA is enabled is not included in this report.  I'll figure it out if I ever have the need.

In the CSV Export, the last column "License Reduction Check" follow a specific criteria:

1: if 'TotalMailboxSize' is less than 50GB, 'In-Place Archive Status' is None, 'LitigationHoldEnabled' is FALSE, and 'Sign In Status' is Blocked, states the following: Passed License Reduction Check (Sign-in is BLOCKED)

2: if 'TotalMailboxSize' is less than 50GB, 'In-Place Archive Status' is None, 'LitigationHoldEnabled' is FALSE, 'Sign In Status' is Enabled, states the following: Passed License Reduction Check (Sign-in is ENABLED)

3: if 'TotalMailboxSize' is more than 50GB, and/or 'In-Place Archive Status' is Active, and/or 'LitigationHoldEnabled' is True, 'Sign In Status' is Enabled, states the following: Failed License Reduction Check (Sign-in is ENABLED) - Please Review

4: if 'TotalMailboxSize' is more than 50GB, and/or 'In-Place Archive Status' is Active, and/or 'LitigationHoldEnabled' is True, 'Sign In Status' is Blocked, states the following: Failed License Reduction Check (Sign-in is BLOCKED) - Please Review
#---------------------------------------------------------------------------------------------------------------------------------------------------------#>

# Connect/Signin to Exchange Online and Azure AD -- sorry but you have to sign into O365/Azure x3 times in order for this script to work properly
#<#
Connect-ExchangeOnline
Connect-AzureAD #needed to pull the license detail
Connect-MsolService #needed to pull Microsoft MFA details
#>

$logdate = get-date -f MM-dd-yyyy
$csvfile = "c:\temp\O365_Audit_Report-$logdate.csv"

# License SKU to friendly name mapping
$LicenseFriendlyNames = @{
    "078d2b04-f1bd-4111-bbd4-b4b1b354cef4" = "Azure Active Directory Premium P1"
	"AAD_PREMIUM" = "Azure Active Directory Premium P1"
	"84a661c4-e949-4bd2-a560-ed7766fcaf2b" = "Azure Active Directory Premium P2"
	"AAD_PREMIUM_P2" = "Azure Active Directory Premium P2"
    "efccb6f7-5641-4e0e-bd10-b4976e1bf68e" = "Enterprise Mobility + Security E3"
    "EMS" = "Enterprise Mobility + Security E3"
    "ee02fd1b-340e-4a4b-b355-4a514e4c8943" = "Exchange Online Archiving for Exchange Online"
	"EXCHANGEARCHIVE_ADDON" = "Exchange Online Archiving for Exchange Online"
    "061f9ace-7d42-4136-88ac-31dc755f143f" = "Intune"
	"INTUNE_A" = "Intune"
    "0c266dff-15dd-4b49-8397-2bb16070ed52" = "Microsoft 365 Audio Conferencing"
	"MCOMEETADV" = "Microsoft 365 Audio Conferencing"
    "dcb1a3ae-b33f-4487-846a-a640262fadf4" = "Microsoft Power Apps Plan 2 Trial"
	"POWERAPPS_VIRAL" = "Microsoft Power Apps Plan 2 Trial"
    "f30db892-07e9-47e9-837c-80727f46fd3d" = "Microsoft Power Automate Free"
	"FLOW_FREE" = "Microsoft Power Automate Free"
    "1f2f344a-700d-42c9-9427-5cea1d5d7ba6" = "Microsoft Stream Trial"
	"STREAM" = "Microsoft Stream Trial"
    "18181a46-0d4e-45cd-891e-60aabd171b4e" = "Office 365 E1"
	"STANDARDPACK" = "Office 365 E1"
    "6fd2c87f-b296-42f0-b197-1e91e994b900" = "Office 365 E3"
	"ENTERPRISEPACK" = "Office 365 E3"
    "c7df2760-2c81-4ef7-b578-5b5392b571df" = "Office 365 E5"
	"ENTERPRISEPREMIUM" = "Office 365 E5"
    "a403ebcc-fae0-4ca2-8c8c-7a907fd6c235" = "Power BI (free)"
	"POWER_BI_STANDARD" = "Power BI (free)"
    "f8a1db68-be16-40ed-86d5-cb42ce701560" = "Power BI Pro"
	"POWER_BI_PRO" = "Power BI Pro"
    "53818b1b-4a27-454b-8896-0dba576410e6" = "Project Plan 3"
	"PROJECTPROFESSIONAL" = "Project Plan 3"
    "4b244418-9658-4451-a2b8-b5e2b364e9bd" = "Visio Plan 1"
	"VISIOONLINE_PLAN1" = "Visio Plan 1"
    "c5928f49-12ba-48f7-ada3-0d743a3601d5" = "Visio Plan 2"
	"VISIOCLIENT" = "Visio Plan 2"
	"bc946dac-7877-4271-b2f7-99d2db13cd2c" = "Dynamics 365 Customer Voice Trial"
	"FORMS_PRO" = "Dynamics 365 Customer Voice Trial"
}

# Group ObjectId for MFA ("RSA MFA Enforced Users" -- this is a dirsynced security group from AD that is tied to an Azure Conditional Policy which handles MFA)
$groupObjectId = "7e9e34e0-8c9d-4b18-baf9-b4b805e5cfd2"

<# --------------------------------------------------------------------------------------------------------------------------------------------------------
#Run script against specific users

# List of user identities
$UserIds = @(
	"user1@contoso.com",
	"user2@contoso.com",
	"user3@contoso.com"
    )

# Get all mailboxes
$Mailboxes = foreach ($UserId in $UserIds) {
    Get-Mailbox -Identity $UserId
}
#-------------------------------------------------------------------------------------------------------------------------------------------------------- #>

#<# --------------------------------------------------------------------------------------------------------------------------------------------------------
#Run Script Against ALL Users
# Get all mailboxes
$Mailboxes = Get-Mailbox -ResultSize Unlimited
#-------------------------------------------------------------------------------------------------------------------------------------------------------- #>

# Initialize error array
$Errors = @()

# Get all users
$users = Get-AzureADUser -All $true -ErrorAction SilentlyContinue

# Add Delegates, In-Place Archive, and LitigationHold information
$Mailboxes | ForEach-Object {
    $User = Get-AzureADUser -ObjectId $_.ExternalDirectoryObjectId -ErrorAction SilentlyContinue
    $Licenses = ($User | Select-Object -ExpandProperty AssignedLicenses | ForEach-Object { $LicenseFriendlyNames[$_.SkuId] }) -join ", "

    $MailboxIdentity = $_.PrimarySmtpAddress

# Get MFA status
    $memberOf = Get-AzureADUserMembership -ObjectId $User.ObjectId | Where-Object {$_.ObjectType -eq "Group"} | Select -ExpandProperty ObjectId
    $mfaStatus = "MFA not Configured"
    if ($memberOf -contains $groupObjectId) {
        $mfaStatus = "RSA MFA Enabled"
    }
    $msolUser = Get-MsolUser -UserPrincipalName $User.UserPrincipalName
    if ($msolUser.StrongAuthenticationRequirements.State -ne $null) {
        if ($mfaStatus -eq "RSA MFA Enabled") {
            $mfaStatus = "RSA & Microsoft MFA Enabled"
        } else {
            $mfaStatus = "Microsoft MFA Enabled"
        }
    }

# Get SendAs permission
$SendAs = (Get-RecipientPermission -Identity $MailboxIdentity | where { -not ($_.Trustee -match "NT AUTHORITY") -and ($_.IsInherited -eq $false)} | ForEach-Object {
    $recipient = Get-Recipient $_.Trustee -ErrorAction SilentlyContinue
    if ($recipient) {
        $recipient.PrimarySmtpAddress
    }
}) -join ";"

# Get Full Access permission
$FullAccess = (Get-MailboxPermission -Identity $MailboxIdentity | where { -not ($_.User -match "NT AUTHORITY") -and ($_.IsInherited -eq $false)} | ForEach-Object {
    $recipient = Get-Recipient $_.User -ErrorAction SilentlyContinue
    if ($recipient) {
        $recipient.PrimarySmtpAddress
    }
}) -join ";"

# Get Send on Behalf permission
$SendOnBehalf = ((Get-Mailbox $MailboxIdentity).GrantSendOnBehalfTo | ForEach-Object {(Get-Recipient $_).PrimarySmtpAddress}) -join ";"

    # In-Place Archive Status
$ArchiveStats = $null
$ArchiveError = $null
Get-MailboxStatistics -Identity $_.Identity -Archive -ErrorAction SilentlyContinue -ErrorVariable ArchiveError | Out-Null

if($ArchiveError) {
    # Write-Host "No archive mailbox found for user: $($_.Identity)"
} else {
    $ArchiveStats = Get-MailboxStatistics -Identity $_.Identity -Archive
}
  
    if ($ArchiveStats) {
        $InPlaceArchiveStatus = "Active"
        $ArchiveSizeBytes = [double]::Parse($ArchiveStats.TotalItemSize.Value.ToString().Split("(")[1].Split(" ")[0])
        if ($ArchiveSizeBytes -ge 1GB) {
            $ArchiveSizeFormatted = "{0:F2} GB" -f ($ArchiveSizeBytes / 1GB)
        } elseif ($ArchiveSizeBytes -ge 1MB) {
            $ArchiveSizeFormatted = "{0:F2} MB" -f ($ArchiveSizeBytes / 1MB)
        } else {
            $ArchiveSizeFormatted = "{0:F2} KB" -f ($ArchiveSizeBytes / 1KB)
        }
    } else {
        $InPlaceArchiveStatus = "None"
        $ArchiveSizeFormatted = $null
    }

# Output
$_ | Select-Object `
DisplayName, 
UserPrincipalName, 
PrimarySmtpAddress, 
@{Name="LastUserActionTime";Expression={(Get-MailboxStatistics $_.PrimarySmtpAddress).LastUserActionTime}}, 
@{Name="LastLogonTime";Expression={(Get-MailboxStatistics $_.PrimarySmtpAddress).LastLogonTime}}, 
@{Name="TotalMailboxSize";Expression={(Get-MailboxStatistics $_.PrimarySmtpAddress).TotalItemSize}}, 
@{Name="User or Shared Mailbox";Expression={(Get-Mailbox $_.PrimarySmtpAddress).RecipientTypeDetails}}, 
@{Name="IsDirSynced";Expression={(Get-Mailbox $_.PrimarySmtpAddress).IsDirSynced}}, 
@{Name="Sign In Status";Expression={if ((Get-Mailbox $_.PrimarySmtpAddress).ExchangeUserAccountControl -eq "none") { "Enabled" } else { "Blocked" }}}, 
@{Name="MFAStatus";Expression={$mfaStatus}},
@{Name="Licenses";Expression={$Licenses}},
@{Name="Full Delegate";Expression={$FullAccess}},
@{Name="SendAs";Expression={$SendAs}},
@{Name="Send on Behalf";Expression={$SendOnBehalf}},
@{Name="In-Place Archive Status";Expression={$InPlaceArchiveStatus}},
@{Name="In-Place Archive Size";Expression={$ArchiveSizeFormatted}},
@{Name="LitigationHoldEnabled";Expression={(Get-Mailbox $_.PrimarySmtpAddress).LitigationHoldEnabled}},
@{Name="License Reduction Check"; Expression={
    $mailboxSize = [double]::Parse((Get-MailboxStatistics $_.PrimarySmtpAddress).TotalItemSize.ToString().Split("(")[1].Split(" ")[0])
    $mailboxSize = $mailboxSize / 1GB
    $litigationHold = (Get-Mailbox $_.PrimarySmtpAddress).LitigationHoldEnabled
    $signInStatus = if ((Get-Mailbox $_.PrimarySmtpAddress).ExchangeUserAccountControl -eq "none") { "Enabled" } else { "Blocked" }

    if($mailboxSize -lt 50 -and $InPlaceArchiveStatus -eq "None" -and $litigationHold -eq $false) {
        if($signInStatus -eq "Blocked") {
            "Passed License Reduction Check (Sign-in is BLOCKED)"
        }
        elseif($signInStatus -eq "Enabled") {
            "Passed License Reduction Check (Sign-in is ENABLED)"
        }
    }
    elseif($mailboxSize -gt 50 -or $InPlaceArchiveStatus -eq "Active" -or $litigationHold -eq $true) {
        if($signInStatus -eq "Blocked") {
            "Failed License Reduction Check (Sign-in is BLOCKED) - Please Review"
        }
        elseif($signInStatus -eq "Enabled") {
            "Failed License Reduction Check (Sign-in is ENABLED) - Please Review"
        }
    }
}}
} | Export-Csv -Path "$csvfile" -NoTypeInformation -Append
