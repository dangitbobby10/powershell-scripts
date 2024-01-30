# Written by: Robert Van Pay -- if you encounter any anomalies please feel free to reach out.
# This script will record a user's current data, export it to CSV, and perform actions to Offboard the user.
#####################################################################################################################################################################
#   ___ _             _    ___                     _   _     _   ___    __       __  __ ___ ____  __ ___    ___   __  __ _                      _ _             ___         _      _   
#  / __| |___ _  _ __| |  / __|___ _ ___ _____ _ _| |_(_)   /_\ |   \  / _|___  |  \/  / __|__ / / /| __|  / _ \ / _|/ _| |__  ___  __ _ _ _ __| (_)_ _  __ _  / __| __ _ _(_)_ __| |_ 
# | (__| / _ \ || / _` | | (__/ _ \ ' \ V / -_) '_|  _|_   / _ \| |) | > _|_ _| | |\/| \__ \|_ \/ _ \__ \ | (_) |  _|  _| '_ \/ _ \/ _` | '_/ _` | | ' \/ _` | \__ \/ _| '_| | '_ \  _|
#  \___|_\___/\_,_\__,_|  \___\___/_||_\_/\___|_|  \__(_) /_/ \_\___/  \_____|  |_|  |_|___/___/\___/___/  \___/|_| |_| |_.__/\___/\__,_|_| \__,_|_|_||_\__, | |___/\__|_| |_| .__/\__|
#                                                                                                                                                       |___/                |_|       
#####################################################################################################################################################################
#  __       ___       ___  __           __   ___  __             ___  __      __   __     __   __     ___  __      __   ___  ___  __   __        __   __          __  
# / _`  /\   |  |__| |__  |__)    |  | /__` |__  |__)    | |\ | |__  /  \    |__) |__) | /  \ |__)     |  /  \    /  \ |__  |__  |__) /  \  /\  |__) |  \ | |\ | / _` 
# \__> /~~\  |  |  | |___ |  \    \__/ .__/ |___ |  \    | | \| |    \__/    |    |  \ | \__/ |  \     |  \__/    \__/ |    |    |__) \__/ /~~\ |  \ |__/ | | \| \__> 
#                                                                                                                                                                
#####################################################################################################################################################################

###############################################################################################################
# Connect to Required MS365 Modules and Import the AD PS Module
###############################################################################################################
#--------------------------------------------------------------------------------------------------------------
    Write-Host "Connecting to MS365 -- you will be asked to log in x3 times." -ForegroundColor Cyan
    Connect-MsolService
    Connect-ExchangeOnline
    Connect-AzureAD
#--------------------------------------------------------------------------------------------------------------
# Import the Active Directory module
    Write-Host "Importing Active Directory Module -- Standby..." -ForegroundColor Cyan
    Import-Module ActiveDirectory
###############################################################################################################
# Create Loop Point. When the script completes - it'll ask if another user needs to be offboareded. By doing 
# it this way, the script will skip reconnecting to the MS365/Azure Modules
###############################################################################################################
    function OffboardUser {
    Remove-Variable * -ErrorAction SilentlyContinue

###############################################################################################################
# Key Defined Variables:
###############################################################################################################
#--------------------------------------------------------------------------------------------------------------
# Define the domain controller to connect to
    $domainController = #"primaryDC.contoso.com"

# Define the server with AADSync
    $AADSyncServer = #"aadsync.contoso.com"

# Define "Disabled Users" OU
    $ou_path = #"OU=Disabled Users,DC=contoso,DC=com"

# Define the path to the "License Friendly Names Script" that transforms MS365 licenses from SKU to Friendly Names (e.g. "ENTERPRISEPACK" = "Office 365 E3")
    $LicenseFriendlyNamesScript = #"C:\"path to..."\LicenseFriendlyNamesScript.ps1"

# Define the 'Date' Variable for the CSV export file
    $date = Get-Date -Format "MM-dd-yyyy"

# Define the path to the CSV file
    #only change the value insde " ". Be sure to keep { } intact as it is used later as a script block IF you have $username in the filepath.
    $csvFilePath = { "c:\users\$env:username\desktop\Offboarding - $username $date.csv" }
#--------------------------------------------------------------------------------------------------------------

###############################################################################################################
# Intial Active Directory Variables and Functions
###############################################################################################################
#--------------------------------------------------------------------------------------------------------------
# Connect to the specific domain controller's Active Directory module
    Write-Host "Connecting to Domain Controller -- Standby..." -ForegroundColor Cyan
    $domainControllerSession = New-PSSession -ComputerName $domainController
    Import-PSSession $domainControllerSession -Module ActiveDirectory -AllowClobber
#--------------------------------------------------------------------------------------------------------------
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.Application]::EnableVisualStyles()

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Enter Offboarding Details'
    # Set a larger initial size for the form
    $form.Size = New-Object System.Drawing.Size(600, 700) 
    $form.StartPosition = 'CenterScreen'

function Add-InputField {
    param($form, $labelText, $topPosition)
    
    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10, $topPosition)
    $label.Size = New-Object System.Drawing.Size(580, 0)  # Adjusted width for the form size
    $label.AutoSize = $true
    $label.Text = $labelText
    $form.Controls.Add($label)
    
# Perform layout to update label size
    $label.PerformLayout()

# Explicitly calculate the Y position for the TextBox
    $textBoxYPosition = $topPosition + $label.Height + 10
    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Location = New-Object System.Drawing.Point(10, $textBoxYPosition)
    $textBox.Size = New-Object System.Drawing.Size(560, 20)
    $form.Controls.Add($textBox)

    return $textBox
    }  

# Add fields using the function, with updates for explicit Y positions
    $usernameBox = Add-InputField $form "Username (e.g., mickey.mouse NOT mickey.mouse@domain.com):" $initialTopPosition
    $forwardingAddressBox = Add-InputField $form 'Forwarding Address:' ($usernameBox.Location.Y + 30)
    $delegateBox1 = Add-InputField $form 'Delegate 1:' ($forwardingAddressBox.Location.Y + 30)
    $delegateBox2 = Add-InputField $form 'Delegate 2:' ($delegateBox1.Location.Y + 30)
    $delegateBox3 = Add-InputField $form 'Delegate 3:' ($delegateBox2.Location.Y + 30)
    $sendasBox1 = Add-InputField $form 'Send As 1:' ($delegateBox3.Location.Y + 30)
    $sendasBox2 = Add-InputField $form 'Send As 2:' ($sendasBox1.Location.Y + 30)
    $sendasBox3 = Add-InputField $form 'Send As 3:' ($sendasBox2.Location.Y + 30)
    $outOfOfficeMessageBox = Add-InputField $form 'Out of Office Message:' ($sendasBox3.Location.Y + 30)

    $submitButton = New-Object System.Windows.Forms.Button
    $submitButton.Text = 'Submit'

# Calculate the Y position for the submit button
    $submitButtonY = $outOfOfficeMessageBox.Location.Y + $outOfOfficeMessageBox.Height + 10
    $submitButton.Location = New-Object System.Drawing.Point(10, $submitButtonY)
    $submitButton.Size = New-Object System.Drawing.Size(560, 23)
    $submitButton.Add_Click({
    
# Collect the input data from the text boxes
    $script:username = $usernameBox.Text
    $script:forwardingAddress = $($forwardingAddressBox.Text)
    $script:delegate1 = $delegateBox1.Text
    $script:delegate2 = $delegateBox2.Text
    $script:delegate3 = $delegateBox3.Text
    $script:sendAs1 = $sendasBox1.Text
    $script:sendAs2 = $sendasBox2.Text
    $script:sendAs3 = $sendasBox3.Text
    $script:outOfOfficeMessage = $outOfOfficeMessageBox.Text

# If the input is valid, close the form and set the form's DialogResult to OK
    $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Close()
    })

    $form.Controls.Add($submitButton)

# Show the form as a dialog and capture the result
    $result = $form.ShowDialog()

# Check the result and perform actions based on it
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        # If the user clicked 'Submit' and the form closed with a DialogResult of OK,               
}
#--------------------------------------------------------------------------------------------------------------
# AD User lookup using the username collected from the form. Will reprompt if inputted user is invalid in Active Directory
    if ($null -eq $username -or $null -eq (Get-ADUser -Identity $username -Server $domainController -ErrorAction SilentlyContinue)) {
                    
    [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
    $PromptText = "ERROR - the username inputted was not found. Please Try again`n`n" +
                "Enter the username of the user being offboarded:`n`n" +
                "Example: mickey.mouse`n`n" +
                "Not Acceptable: mickey.mouse@domain.com, mickey.mouse@domain.local, or DOMAIN\mickey.mouse"
    $Title = "Input the Username of the User Being Offboarded"
    $username = $null

    while ($null -eq $username -or $null -eq (Get-ADUser -Identity $username -Server $domainController -ErrorAction SilentlyContinue)) {
        $username = [Microsoft.VisualBasic.Interaction]::InputBox($PromptText, $Title)
        
        if ([string]::IsNullOrWhiteSpace($username)) {
            Write-Host "No username entered. Exiting script." -ForegroundColor Magenta
            Exit
        }

        if ($null -eq (Get-ADUser -Identity $username -Server $domainController -ErrorAction SilentlyContinue)) {
            [System.Windows.Forms.MessageBox]::Show("$username was not found in Active Directory. Please try again.", "User Not Found")
            }
        }
    }
#--------------------------------------------------------------------------------------------------------------
# Get the user's current AD information ready for CSV Export
    Write-Host "Preparing CSV Export of User's Current Configs -- Standby..." -ForegroundColor Cyan
    $user = Get-ADUser -Identity $username -Properties Description, IPPhone, MemberOf, mail, UserPrincipalName, DistinguishedName -Server $domainController
    $description = $user.Description
    $ipPhone = $user.IPPhone
    $email = $user.mail
    $upn = $user.UserPrincipalName
    $groups = $user.MemberOf | ForEach-Object {(Get-ADGroup -Identity $_).Name} # Safeguard against null values
    if ([string]::IsNullOrEmpty($ipPhone)) {
        $ipPhone = "No IP Phone Recorded at time of Offboarding"
    }
    $groupList = $groups -join ";"
#--------------------------------------------------------------------------------------------------------------
# Function to extract the OU from the DistinguishedName
    function Get-OUFromDN($distinguishedName) {
        return ($distinguishedName -split '(?<=,)', 2)[1] -replace 'DC=', '' -replace ',', '.'
        }
    $ou = Get-OUFromDN($user.DistinguishedName)
#--------------------------------------------------------------------------------------------------------------
# More Defining Variables
    $mailbox = Get-Mailbox -Identity $upn
    $mailboxStats = Get-MailboxStatistics -Identity $upn
    $AADUsername = Get-AzureADUser -Filter "UserPrincipalName eq '$upn'"
#--------------------------------------------------------------------------------------------------------------

###############################################################################################################
# MS365/AAD Variables and Functions
###############################################################################################################
#--------------------------------------------------------------------------------------------------------------
# Create hashtable to store the upcoming properties
    $properties = @{}
#--------------------------------------------------------------------------------------------------------------
# Import 'LicenseFriendlyNamesScript' for the MS365 Licenses. Reads as the actual license rather than the SKU.
    . $LicenseFriendlyNamesScript

# Join friendly license names and add to properties hashtable
    $properties['Licenses'] = $friendlyLicenseNames -join ", "
#--------------------------------------------------------------------------------------------------------------
# Check if mailbox is at or over 50GB
    $mailboxSizeValue = $mailboxStats.TotalItemSize.ToString()
    $isInGB = $mailboxSizeValue -match 'GB'
    $isMailboxOver50GB = $false

    if ($isInGB) {
        $mailboxSizeGB = [double]::Parse($mailboxSizeValue.Split(" ")[0])
        $isMailboxOver50GB = $mailboxSizeGB -ge 50
    }
#--------------------------------------------------------------------------------------------------------------
# Check if in-place archive is enabled
    $isInPlaceArchiveEnabled = $mailbox.ArchiveStatus -eq "Active"

    <# ►►Not needed anymore - converting to shared mailbox regardless of Mailbox Size/In-Place Archive Status. Below was my attempt to get it to work but I couldn't

    # Check if Mailbox is over 50GB and/or has In-Place Archive Enabled. Export CSV whether the account will be converted or not
    if ($isMailboxOver50GB -and $isInPlaceArchiveEnabled) {
        $properties["Mailbox Conversion to Shared"] = "Mailbox is too Large & has In-Place Archive Enabled. Cannot Convert"
    } elseif ($isMailboxOver50GB) {
        $properties["Mailbox Conversion to Shared"] = "Mailbox is too large to Convert"
    } elseif ($isInPlaceArchiveEnabled) {
        $properties["Mailbox Conversion to Shared"] = "Mailbox has In-Place Archive Enabled, Cannot Convert"
    } else {
        $properties["Mailbox Conversion to Shared"] = "Mailbox will be Converted to Shared"
    }
    #>
#--------------------------------------------------------------------------------------------------------------
# Get all AD group membership(s)
    $UserGroups = Get-AzureADUser -ObjectId $upn
    $Memberships = Get-AzureADUserMembership -ObjectId $UserGroups.ObjectId | Where-object { $_.ObjectType -eq "Group" }
    $groupNames = $Memberships | Select-Object -ExpandProperty DisplayName
    $properties['MS365 Groups'] = $groupNames -join ", "
#--------------------------------------------------------------------------------------------------------------
# Get all the mailbox's O365 licenses
    $licenses = (Get-MsolUser -UserPrincipalName $upn).Licenses

# Convert license SKUs to friendly names
    $friendlyLicenseNames = @()
    foreach ($license in $licenses) {
        $skuId = $license.AccountSkuId
        $sku = $skuId.Split(":")[1]

        $friendlyName = $LicenseFriendlyNames[$sku]
        if (-not $friendlyName) {
            $friendlyName = $sku
        }

        $friendlyLicenseNames += $friendlyName
    }
# Assign licenses to properties after populating friendlyLicenseNames
    $properties['Licenses'] = $friendlyLicenseNames -join ", "
#--------------------------------------------------------------------------------------------------------------
# Check if forwarding is enabled
    $isForwardingEnabled = $mailbox.ForwardingSmtpAddress -ne $null
    $currentforwardingAddress = $mailbox.ForwardingSmtpAddress

# If there is forwarding, record the forwarding account
    if ($isForwardingEnabled) {
        $properties["Forwarding To"] = $currentforwardingAddress
    } else {
        $properties["Forwarding To"] = "No Forwarding Prior to Offboarding"
    }
#--------------------------------------------------------------------------------------------------------------
# Check for Delegates
    $delegates = Get-MailboxPermission -Identity $upn | Where-Object {$_.AccessRights -eq "FullAccess" -and $_.IsInherited -eq $false} | Select-Object -ExpandProperty User
    $properties['Delegates'] = $delegates -join ", "
#--------------------------------------------------------------------------------------------------------------
# Check for SendAs
    $sendAs = Get-RecipientPermission -Identity $upn |
            Where-Object { $_.AccessRights -eq "SendAs" } |
            Where-Object { $_.Trustee -ne "NT AUTHORITY\SELF" } |
            Select-Object -ExpandProperty Trustee
    $properties['SendAs'] = $sendAs -join ", "
#--------------------------------------------------------------------------------------------------------------
# Check for SendonBehalf
    $sendOnBehalf = $mailbox.GrantSendOnBehalfTo | ForEach-Object { (Get-Recipient $_).DisplayName }
    $properties['SendOnBehalf'] = $sendOnBehalf -join ", "
#--------------------------------------------------------------------------------------------------------------

#####################################################################################################################################################################
#  ___      __   __   __  ___     __       ___         ___  __      __   __       
# |__  \_/ |__) /  \ |__)  |     |  \  /\   |   /\      |  /  \    /  ` /__` \  / 
# |___ / \ |    \__/ |  \  |     |__/ /~~\  |  /~~\     |  \__/    \__, .__/  \/  
#                                                                                
#####################################################################################################################################################################
#--------------------------------------------------------------------------------------------------------------
    $finalResult = New-Object PSObject
    $finalResult | Add-Member -MemberType NoteProperty -Name "AD Username" -Value $username
    $finalResult | Add-Member -MemberType NoteProperty -Name "AD Description" -Value $description
    $finalResult | Add-Member -MemberType NoteProperty -Name "AD Organizational Unit" -Value $ou
    $finalResult | Add-Member -MemberType NoteProperty -Name "AD Email" -Value $email
    $finalResult | Add-Member -MemberType NoteProperty -Name "AD UPN" -Value $upn
    $finalResult | Add-Member -MemberType NoteProperty -Name "AD IPPhone" -Value $ipPhone
    $finalResult | Add-Member -MemberType NoteProperty -Name "AD Groups" -Value $groupList
    $finalResult | Add-Member -MemberType NoteProperty -Name "MS365 Groups" -Value $properties['MS365 Groups']
    $finalResult | Add-Member -MemberType NoteProperty -Name "Forwarding To" -Value $properties['Forwarding To']
    $finalResult | Add-Member -MemberType NoteProperty -Name "Delegates" -Value $properties['Delegates']
    $finalResult | Add-Member -MemberType NoteProperty -Name "SendAs" -Value $properties['SendAs']
    $finalResult | Add-Member -MemberType NoteProperty -Name "SendOnBehalf" -Value $properties['SendOnBehalf']
    $finalResult | Add-Member -MemberType NoteProperty -Name "Licenses" -Value $properties['Licenses']
#--------------------------------------------------------------------------------------------------------------
# Export result to CSV file
    $csvValue = &$csvFilePath #script block mentioned earlier in the script
    $finalResult | Export-Csv -Path $csvValue -Append -NoTypeInformation
    Write-Host "$username's AD and MS365 data has been recorded to '$csvValue' -- the script will now being the offboarding process" -ForegroundColor Green
#--------------------------------------------------------------------------------------------------------------

#####################################################################################################################################################################
#  __   ___  __             __   ___  ___  __   __        __   __          __  
# |__) |__  / _` | |\ |    /  \ |__  |__  |__) /  \  /\  |__) |  \ | |\ | / _` 
# |__) |___ \__> | | \|    \__/ |    |    |__) \__/ /~~\ |  \ |__/ | | \| \__> 
#                                                                              
#####################################################################################################################################################################
#       __  ___         ___     __     __   ___  __  ___  __   __      
#  /\  /  `  |  | \  / |__     |  \ | |__) |__  /  `  |  /  \ |__) \ / 
# /~~\ \__,  |  |  \/  |___    |__/ | |  \ |___ \__,  |  \__/ |  \  |  
#                                                                      
#####################################################################################################################################################################
#--------------------------------------------------------------------------------------------------------------
# Disable the AD account
    Disable-ADAccount -Identity $username -Server $domainController
    Write-Host "$username's AD Account has been Disabled" -ForegroundColor Green
#--------------------------------------------------------------------------------------------------------------
# Reset the AD account's password to something random (21 complex unique characters)
    $specialCharacters = "~!@#$%^&*"
    $password = -join ((48..57) + (65..90) + (97..122) + [int[]][char[]]$specialCharacters | Get-Random -Count 21 | % {[char]$_})
    Set-ADAccountPassword -Identity $username -NewPassword (ConvertTo-SecureString -AsPlainText $password -Force) -Server $domainController
    Write-Host "$username's password has been set to:" -ForegroundColor Green
    Write-Host "$password" -ForegroundColor Yellow
#--------------------------------------------------------------------------------------------------------------
# Change the AD Description Field to state "Disabled on (current date)"
    Set-ADUser -Identity $username -Description "Disabled on $date" -Server $domainController

# Refresh the $user object to get the updated description
    $user = Get-ADUser -Identity $username -Properties Description

# Display the updated description
    $newdescription = $user.Description
    Write-Host "$username's AD Description has been updated to:" -ForegroundColor Green
    Write-Host "$newdescription" -ForegroundColor Yellow
#--------------------------------------------------------------------------------------------------------------
# Remove the "IP Phone" entry
    Set-ADUser -Identity $username -Clear IPPhone -Server $domainController
    Write-Host "$username's IP Phone Entry has been cleared" -ForegroundColor Green
#--------------------------------------------------------------------------------------------------------------
# Remove all groups EXCEPT for "Domain Users"
    $groups | Where-Object {$_ -ne "Domain Users"} | ForEach-Object {Remove-ADGroupMember -Identity $_ -Members $username -Confirm:$false}
    Write-Host "$username's AD groups have been removed" -ForegroundColor Green
#--------------------------------------------------------------------------------------------------------------
# Hide from the Global Address List
    Set-ADUser -Identity $username –Replace @{msExchHideFromAddressLists=$true} –Server $domainController
    Write-Host "$username's account has been hidden from the GAL" -ForegroundColor Green
#--------------------------------------------------------------------------------------------------------------
# Move User to Disabled Users (or specified) OU
    try {
    # Attempt to retrieve the user from AD
        $user = Get-ADUser -Identity $username -ErrorAction Stop

    # Move the user to the specified OU
        Move-ADObject -Identity $user.DistinguishedName -TargetPath $ou_path -ErrorAction Stop
        Write-Host "$username's AD Account has been moved to '$ou'" -ForegroundColor Green
    }
    
    catch {
    # Handle errors, such as user not found or lack of permissions
        Write-Host "Error: Unable to move $username to '$ou'. Details: $_" -ForegroundColor Red
    }
#--------------------------------------------------------------------------------------------------------------
# Update User's AD DisplayName

# Get the user's current display name and full name
    $displayName = $user.DisplayName
    $fullName = $user.Name

# Append "Offboarded - " at the beginning of the display name
    $newDisplayName = "Offboarded - $fullName"

# Set the new display name for the user's AD object
    Set-ADUser -Identity $username -DisplayName $newDisplayName -Server $domainController
    Write-Host "$username's AD Display Name has been updated to:" -ForegroundColor Green
    Write-Host "$newDisplayName" -ForegroundColor Yellow
#--------------------------------------------------------------------------------------------------------------

###############################################################################################################
# Run AADConnect Sync (Delta Sync) and wait for 2 minutes
###############################################################################################################
#--------------------------------------------------------------------------------------------------------------
# Run a Delta Sync to sync the changes to MS365
    Write-Host "AADSync Command has been Executed - Standby..." -ForegroundColor Cyan
    Invoke-Command –ComputerName $AADSyncServer –ScriptBlock {Start-ADSyncSyncCycle –PolicyType Delta}

# Wait 2 Minutes to Allow AADSync to be performed
    $delay = 120
    $Counter_Form = New-Object System.Windows.Forms.Form
    $Counter_Form.Text = "Waiting for 2 Minutes for AADSync"
    $Counter_Form.Width = 450
    $Counter_Form.Height = 200
    $Counter_Label = New-Object System.Windows.Forms.Label
    $Counter_Label.AutoSize = $true
    $Counter_Label.ForeColor = "Green"
    $normalfont = New-Object System.Drawing.Font("Times New Roman",14)
    $Counter_Label.Font = $normalfont
    $Counter_Label.Left = 20
    $Counter_Label.Top = 20
    $Counter_Form.Controls.Add($Counter_Label)
    while ($delay -ge 0)
    {
    $Counter_Form.Show()
    $Counter_Label.Text = "Seconds Remaining: $($delay)"
    if ($delay -lt 5)
    { 
        $Counter_Label.ForeColor = "Red"
        $fontsize = 20-$delay
        $warningfont = New-Object System.Drawing.Font("Times New Roman",$fontsize,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold -bor [System.Drawing.FontStyle]::Underline))
        $Counter_Label.Font = $warningfont
    } 
    start-sleep 1
    $delay -= 1
    }
    $Counter_Form.Close()
    Write-Host "AAD Sync to MS365 has been ran" -ForegroundColor Green
#--------------------------------------------------------------------------------------------------------------

#####################################################################################################################################################################
#       __    /      __       __   ___ 
# |\/| /__`  /   /\   / |  | |__) |__  
# |  | .__/ /   /~~\ /_ \__/ |  \ |___ 
#####################################################################################################################################################################
Write-Host "ATTENTION: At this part of the script, it will restore the mailbox as a cloud account and then clear the immutable ID." -ForegroundColor Cyan
Write-Host "This takes a while, we are at the mercy of Microsoft. Go grab a coffee or eat lunch if you haven't." -ForegroundColor Cyan

# Step 1: Check if the mailbox is active and wait until it's not found
do {
    $mailboxping = Get-Mailbox -Identity $upn -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 10
	} while ($mailboxping)

# Step 2: Verify the mailbox is in the 'Soft Deleted' mailboxes
	$mailboxDeleted = Get-Mailbox -SoftDeletedMailbox -Identity $upn -ErrorAction SilentlyContinue

if ($mailboxDeleted) {
    # Step 3: Restore the mailbox
    Write-Host "$upn found in MS' Deleted Users. Restoring $upn -- Standby..." -ForegroundColor Cyan
    Restore-MsolUser -UserPrincipalName $upn
    
    # Verify if the mailbox is restored
    do {
        $restoredMailbox = Get-Mailbox -Identity $upn -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 10
		} while (-not $restoredMailbox)

    # Step 4: Clear the immutable ID
		Set-Mailbox -Identity $upn -ImmutableId $null
		Write-Host "Mailbox restored and Immutable ID cleared for $upn" -ForegroundColor Cyan
		
		} else {
			Write-Host "Mailbox is not in the soft deleted mailboxes." -ForegroundColor Magenta
				}

#--------------------------------------------------------------------------------------------------------------
# Set sign-in status: Block the sign-in status.
    Set-AzureADUser -ObjectID $upn -AccountEnabled $false
    write-host "MS365 Sign-in as been Blocked" -ForegroundColor Green
#--------------------------------------------------------------------------------------------------------------
# Disconnect existing sessions: Terminate any existing user sessions.
    $AzureObjectID = Get-AzureADUser -ObjectId $upn
    $objectId = $AzureObjectID.ObjectId

    try {
        Revoke-AzureADUserAllRefreshToken -ObjectId $objectID
        write-host "MS365 & Azure Sessions have been Revoked & Disconnected" -ForegroundColor Green
    }
    catch {
        Write-Host "Error revoking refresh tokens: $_" -ForegroundColor Magenta
    }
#--------------------------------------------------------------------------------------------------------------

###############################################################################################################
# Convert to Shared Mailbox and wait for 2 minutes
###############################################################################################################
#--------------------------------------------------------------------------------------------------------------
Write-Host "Mailbox is being converted to a Shared Mailbox - Standby..." -ForegroundColor Cyan

# Convert to shared mailbox and verify after conversion
    Set-Mailbox -Identity $upn -Type Shared
	$delay = 120
	$Counter_Form = New-Object System.Windows.Forms.Form
	$Counter_Form.Text = "Waiting for 2 Minutes for Mailbox to Convert to Shared"
	$Counter_Form.Width = 450
	$Counter_Form.Height = 200
	$Counter_Label = New-Object System.Windows.Forms.Label
	$Counter_Label.AutoSize = $true
	$Counter_Label.ForeColor = "Green"
	$normalfont = New-Object System.Drawing.Font("Times New Roman",14)
	$Counter_Label.Font = $normalfont
	$Counter_Label.Left = 20
	$Counter_Label.Top = 20
	$Counter_Form.Controls.Add($Counter_Label)
	while ($delay -ge 0)
	{
	  $Counter_Form.Show()
	  $Counter_Label.Text = "Seconds Remaining: $($delay)"
	  if ($delay -lt 5)
	  { 
		 $Counter_Label.ForeColor = "Red"
		 $fontsize = 20-$delay
		 $warningfont = New-Object System.Drawing.Font("Times New Roman",$fontsize,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold -bor [System.Drawing.FontStyle]::Underline))
		 $Counter_Label.Font = $warningfont
	  } 
	 start-sleep 1
	 $delay -= 1
	}
	$Counter_Form.Close()
#--------------------------------------------------------------------------------------------------------------

#################################################################################################################################################################
# Remove MS365 Licenses. Performs 3 checks and 4 actions based on the checks:
# 1: Checks if Mailbox is, or greater than 50GB and In-Place Archive is enabled -- Removes all licenses except for "E3 and E5"
# 2: Checks if Mailbox is, or greater than 50GB and In-Place Archive is disabled -- Removes all licenses except for "E3 and E5"
# 3: Checks In-Place Archive is enabled and if Mailbox is less than 50GB -- Removes all licenses except for "E3, E5 and 'Exchange Online Archiving for Exchange'"
# 4: If the first 3 checks are not met -- Removes all licenses
#################################################################################################################################################################
#--------------------------------------------------------------------------------------------------------------
# Check if mailbox is, or greater than 50GB and In-Place Archive is enabled: Removes all licenses except for "E3 and E5"
    if ($isMailboxOver50GB -and $isInPlaceArchiveEnabled){
        # Strip all O365 licenses except for the E3 and E5 licenses
        $ExcludedLicenses = @("ENTERPRISEPACK", "ENTERPRISEPREMIUM") #Office 365 E3, & E5
        $AssignedLicensesTable = Get-AzureADUser -ObjectId $upn | Get-AzureADUserLicenseDetail | Select-Object @{n = "License"; e = { $_.SkuPartNumber } }, skuid
        if ($AssignedLicensesTable) {
            $licensesToRemove = @()
            foreach ($license in $AssignedLicensesTable) {
                if ($license.License -notin $ExcludedLicenses) {
                    $licensesToRemove += $license.skuid
                }
            }

            if ($licensesToRemove.Count -gt 0) {
                $body = @{
                    addLicenses    = @()
                    removeLicenses = $licensesToRemove
                }
                Set-AzureADUserLicense -ObjectId $upn -AssignedLicenses $body
                Write-Host "$username's Mailbox was converted to Shared and all licenses except for E3, and/or E5 have been removed. The Mailbox is larger than 50GB and there is an In-Place Archive." -ForegroundColor Green
            } else {
                Write-Host "$username's Mailbox was converted to Shared and no licenses were removed because only E3 and/or E5 license(s) are assigned. The Mailbox is larger than 50GB and there is an In-Place Archive." -ForegroundColor Cyan
                }
            }
    }
#--------------------------------------------------------------------------------------------------------------
# Check if mailbox is or greater than 50GB: Remove all licenses except for "E3 and E5"
    elseif ($isMailboxOver50GB -and -not $isInPlaceArchiveEnabled){
        # Strip all O365 licenses except for the E3 and E5 licenses
        $ExcludedLicenses = @("ENTERPRISEPACK", "ENTERPRISEPREMIUM") #Office 365 E3, & E5
        $AssignedLicensesTable = Get-AzureADUser -ObjectId $upn | Get-AzureADUserLicenseDetail | Select-Object @{n = "License"; e = { $_.SkuPartNumber } }, skuid
        if ($AssignedLicensesTable) {
            $licensesToRemove = @()
            foreach ($license in $AssignedLicensesTable) {
                if ($license.License -notin $ExcludedLicenses) {
                    $licensesToRemove += $license.skuid
                }
            }

            if ($licensesToRemove.Count -gt 0) {
                $body = @{
                    addLicenses    = @()
                    removeLicenses = $licensesToRemove
                }
                Set-AzureADUserLicense -ObjectId $upn -AssignedLicenses $body
                Write-Host "$username's Mailbox was converted to Shared and all licenses except for E3, and/or E5 have been removed. The Mailbox is larger than 50GB" -ForegroundColor Green
            } else {
                Write-Host "$username's Mailbox was converted to Shared and no licenses were removed because only E3 and/or E5 License(s) are assigned. The Mailbox is larger than 50GB" -ForegroundColor Cyan
                }
            }
    }
#--------------------------------------------------------------------------------------------------------------
# Check if mailbox In-Place Archive is enabled
    elseif ($isInPlaceArchiveEnabled -and -not $isMailboxOver50GB){
        # Strip all O365 licenses except for the E3, E5, and 'Exchange Online Archiving for Exchange' Online License(s)
        $ExcludedLicenses = @("ENTERPRISEPACK", "ENTERPRISEPREMIUM", "EXCHANGEARCHIVE_ADDON") #Office 365 E3, E5, & 'Exchange Online Archiving for Exchange Online'
        $AssignedLicensesTable = Get-AzureADUser -ObjectId $upn | Get-AzureADUserLicenseDetail | Select-Object @{n = "License"; e = { $_.SkuPartNumber } }, skuid
        if ($AssignedLicensesTable) {
            $licensesToRemove = @()
            foreach ($license in $AssignedLicensesTable) {
                if ($license.License -notin $ExcludedLicenses) {
                    $licensesToRemove += $license.skuid
                }
            }

            if ($licensesToRemove.Count -gt 0) {
                $body = @{
                    addLicenses    = @()
                    removeLicenses = $licensesToRemove
                }
                Set-AzureADUserLicense -ObjectId $upn -AssignedLicenses $body
                Write-Host "$username's Mailbox was converted to Shared and all licenses except for E3, E5, and/or 'Exchange Online Archiving for Exchange Online' have been removed. The Mailbox has an In-Place Archive enabled" -ForegroundColor Green
            } else {
                Write-Host "$username's Mailbox was converted to Shared and no licenses were removed because only the E3, E5, and/or 'Exchange Online Archiving for Exchange Online' License(s) were assigned. The Mailbox has an In-Place Archive enabled" -ForegroundColor Cyan
                }
            }
        }
#--------------------------------------------------------------------------------------------------------------
# If Mailbox is less than 50GB and In-Place Archive is not enabled, remove all licenses
    else {
            # Strip all O365 licenses
            $AssignedLicensesTable = Get-AzureADUser -ObjectId $upn | Get-AzureADUserLicenseDetail | Select-Object @{n = "License"; e = { $_.SkuPartNumber } }, skuid
            if ($AssignedLicensesTable) {
                $body = @{
                    addLicenses    = @()
                    removeLicenses = @($AssignedLicensesTable.skuid)
                }
                Set-AzureADUserLicense -ObjectId $upn -AssignedLicenses $body
                write-host "$username's Mailbox was converted to Shared and all Licenses have been removed" -ForegroundColor Green
            }
        }
#--------------------------------------------------------------------------------------------------------------

###############################################################################################################
# Configure Email Forwarding
###############################################################################################################
#--------------------------------------------------------------------------------------------------------------
# Configure Email Forwarding
    [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
    Function Get-ForwardingAddress {
    # Function to validate email address format and existence
    Function Validate-EmailAddress($address) {
        if ($address -match "^[^@]+@[^@]+\.[^@]+$") {
            try {
                $resolvedMailbox = Get-Recipient $address
                if ($resolvedMailbox -ne $null) {
                    return $resolvedMailbox.PrimarySmtpAddress
                } else {
                    Write-Host "No mailbox found for the provided address. Please try again." -ForegroundColor Yellow
                    return $null
                }
            } catch {
                Write-Host "Failed to resolve email address: $_" -ForegroundColor Red
                return $null
            }
        } else {
            Write-Host "Invalid email format. Please try again." -ForegroundColor Yellow
            return $null
        }
    }

# Check if $forwardingAddress is already set and valid
    if ($forwardingAddress) {
        $validatedAddress = Validate-EmailAddress $forwardingAddress
        if ($validatedAddress) {
            return $validatedAddress
        }
    }

# Prompt for new forwarding address
    do {
        $PromptText = "Forwarding address ($forwardingAddress) not found in MS365. Please try again.`n`nExample: user@domain.com`n`nClick Cancel if no forwarding needs to be applied."
        $Title = "Forwarding Email Entry"
        $forwardingAddressInput = [Microsoft.VisualBasic.Interaction]::InputBox($PromptText, $Title)

        if ([string]::IsNullOrEmpty($forwardingAddressInput)) {
            Write-Host "No forwarding address provided. Moving to the next step." -ForegroundColor Yellow
            return $null
            }

            $validatedAddress = Validate-EmailAddress $forwardingAddressInput
        } while (-not $validatedAddress)

        return $validatedAddress
    }

    $forwardingAddress = Get-ForwardingAddress

# Check if a forwarding address was provided
    if ($forwardingAddress -ne $null) {
        # Update forwarding address
        Set-Mailbox -Identity $upn -ForwardingSmtpAddress $forwardingAddress -DeliverToMailboxAndForward $true
        Write-Host "$upn's emails will also forward to $forwardingAddress" -ForegroundColor Green
    } else {
        Write-Host "No forwarding address set for $upn." -ForegroundColor Yellow
        }
#--------------------------------------------------------------------------------------------------------------
# Configure Delegate Permissions
    # Function to validate delegate address in MS365
    Function Validate-Delegate($delegateAddress) {
        if ($delegateAddress -match "^[^@]+@[^@]+\.[^@]+$") {
            try {
                # Replace 'Get-Mailbox' with the appropriate cmdlet or command for your MS365 environment
                $mailbox = Get-Mailbox -Identity $delegateAddress -ErrorAction Stop
                return $delegateAddress
            } catch {
                Write-Host "Delegate address ($delegateAddress) not found in MS365. Please try again." -ForegroundColor Yellow
                return $null
            }
        } else {
            Write-Host "Invalid email format for address ($delegateAddress). Please try again." -ForegroundColor Yellow
            return $null
        }
    }

# Configure Delegate Permissions
    $delegateAddresses = @($delegate1, $delegate2, $delegate3)

    foreach ($delegate in $delegateAddresses) {
        if ($delegate -ne "") {
            $validatedDelegate = $null
            while ($null -eq $validatedDelegate) {
                $validatedDelegate = Validate-Delegate $delegate
                if ($null -eq $validatedDelegate) {
                    # Prompt for re-entry
                    $delegatePrompt = "Enter a valid delegate email address in MS365 for address ($delegate):"
                    $delegate = [Microsoft.VisualBasic.Interaction]::InputBox($delegatePrompt, "Delegate Email Entry")
                    if ([string]::IsNullOrWhiteSpace($delegate)) {
                        Write-Host "Entry for address ($delegate) canceled. Moving to the next delegate." -ForegroundColor Yellow
                        break
                    }
                }
            }

# Once validated, add permissions
            if ($validatedDelegate -ne $null) {
                Add-MailboxPermission -Identity $upn -User $validatedDelegate -AccessRights FullAccess -InheritanceType All
            }
        }
    }
#--------------------------------------------------------------------------------------------------------------
# Configure Send-As Permissions
    # Function to validate Send-As address in MS365
    Function Validate-SendAs($sendAsAddress) {
        if ($sendAsAddress -match "^[^@]+@[^@]+\.[^@]+$") {
            try {
                # Replace 'Get-Mailbox' with the appropriate cmdlet or command for your MS365 environment
                $mailbox = Get-Mailbox -Identity $sendAsAddress -ErrorAction Stop
                return $sendAsAddress
            } catch {
                Write-Host "Send-As address ($sendAsAddress) not found in MS365. Please try again." -ForegroundColor Yellow
                return $null
            }
        } else {
            Write-Host "Invalid email format for address ($sendAsAddress). Please try again." -ForegroundColor Yellow
            return $null
        }
    }

    # Configure Send-As Permissions
    $sendAsAddresses = @($sendAs1, $sendAs2, $sendAs3)

    foreach ($sendAs in $sendAsAddresses) {
        if ($sendAs -ne "") {
            $validatedSendAs = $null
            while ($null -eq $validatedSendAs) {
                $validatedSendAs = Validate-SendAs $sendAs
                if ($null -eq $validatedSendAs) {
                    # Prompt for re-entry
                    $sendAsPrompt = "Enter a valid Send-As email address in MS365 for address ($sendAs):"
                    $sendAs = [Microsoft.VisualBasic.Interaction]::InputBox($sendAsPrompt, "Send-As Email Entry")
                    if ([string]::IsNullOrWhiteSpace($sendAs)) {
                        Write-Host "Entry for address ($sendAs) canceled. Moving to the next Send-As address." -ForegroundColor Yellow
                        break
                    }
                }
            }
            # Once validated, add Send-As permissions
            if ($validatedSendAs -ne $null) {
                Add-RecipientPermission -Identity $upn -Trustee $validatedSendAs -AccessRights SendAs -Confirm:$false
            }
        }
    }
#--------------------------------------------------------------------------------------------------------------

###############################################################################################################
# Configure out of office reply
###############################################################################################################
if ($outOfOfficeMessage -ne "") {
    Set-MailboxAutoReplyConfiguration -Identity $upn -AutoReplyState Enabled -ExternalAudience All -ExternalMessage $outOfOfficeMessage -InternalMessage $outOfOfficeMessage 
    Write-Host "The Out-Of-Office-Rely intputted was successfully applied." -ForegroundColor Green
} else {
    Write-Host "No Out of Office message set for $upn." -ForegroundColor Yellow
}
#--------------------------------------------------------------------------------------------------------------


###############################################################################################################
# Remove user from all MS365 and Security Groups
###############################################################################################################
#--------------------------------------------------------------------------------------------------------------
# Get user details
    $UserDetails = Get-AzureADUser -ObjectId $upn
    $UserId = $UserDetails.ObjectId

# Removing All Groups (MS365, Security, & Distribution)
    $Memberships = Get-AzureADUserMembership -ObjectId $UserId | Where-object { $_.ObjectType -eq "Group"}

# Remove user from all MS365 and Security Groups
    foreach ($group in $Memberships) {
        try {
            Remove-AzureADGroupMember -ObjectId $group.ObjectId -MemberId $UserId
            Write-Host "Successfully removed user $($UserDetails.DisplayName) from group $($group.DisplayName)" -ForegroundColor Green
        }
        catch {
            Write-Host "Error removing user $($UserDetails.DisplayName) from group $($group.DisplayName)" -ForegroundColor Magenta
            Write-Host "Note: Some groups that are unable to be removed may be ADSynced or a Distribution Group, which the next command will catch." -ForegroundColor Cyan
            Write-Host "Also Note: There are a handful of groups that are applied at the Organization level that cannot be removed." -ForegroundColor Cyan
        }
    }

# Remove user from all Distrubtion Groups
    $DistinguishedName = $Mailbox.DistinguishedName 
    Get-DistributionGroup -Filter "Members -eq '$DistinguishedName'" | ForEach-Object {
        try {
            Remove-DistributionGroupMember -Identity $_.Identity -Member $DistinguishedName -Confirm:$false
            Write-Host "Successfully removed user $($UserDetails.DisplayName) from distribution group $($_.DisplayName)" -ForegroundColor Green
        }
        catch {
            Write-Host "Error removing user $($UserDetails.DisplayName) from distribution group $($_.DisplayName)" -ForegroundColor Magenta
        }
    }
#--------------------------------------------------------------------------------------------------------------------------------------------------------

Write-Host "The Offboarding Script has been fully ran" -ForegroundColor Green

}

###############################################################################################################
# This part of the script will prompt if another user needs to be offboarded. If yes, the script to 
# execute again but skips the step that connects to the MS365/Azure Powershell Modules
###############################################################################################################
function Show-Prompt {
    # Create a new form
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'User Offboarding'
    $form.Size = New-Object System.Drawing.Size(300,200)
    $form.StartPosition = 'CenterScreen'

    # Add a label with your text
    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10,20)
    $label.Size = New-Object System.Drawing.Size(280,20)
    $label.Text = 'Do you need to Offboard another user?'
    $form.Controls.Add($label)

    # Create a "Yes" button
    $yesButton = New-Object System.Windows.Forms.Button
    $yesButton.Location = New-Object System.Drawing.Point(50,100)
    $yesButton.Size = New-Object System.Drawing.Size(75,23)
    $yesButton.Text = 'Yes'
    $yesButton.DialogResult = [System.Windows.Forms.DialogResult]::Yes
    $form.AcceptButton = $yesButton
    $form.Controls.Add($yesButton)

    # Create a "No" button
    $noButton = New-Object System.Windows.Forms.Button
    $noButton.Location = New-Object System.Drawing.Point(150,100)
    $noButton.Size = New-Object System.Drawing.Size(75,23)
    $noButton.Text = 'No'
    $noButton.DialogResult = [System.Windows.Forms.DialogResult]::No
    $form.CancelButton = $noButton
    $form.Controls.Add($noButton)

    # Show the form
    return $form.ShowDialog()
}

# Initial call to the offboarding function
    OffboardUser

# Loop to show the prompt and repeat the task if "Yes" is selected
    do {
        $result = Show-Prompt
        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
            OffboardUser
        }
    } while ($result -eq [System.Windows.Forms.DialogResult]::Yes)

# Script ends when "No" is clicked