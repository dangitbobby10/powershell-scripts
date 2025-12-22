# PowerShell Script for Interactive Mailbox Permission Management
# This script allows you to view and modify mailbox permissions (Full Access, SendAs, SendOnBehalf) for any mailbox

#Connect to Exchange Online
Connect-ExchangeOnline

# Function to display Full Access (Delegate) permissions
function Show-FullAccessPermissions {
    param($Mailbox)
    
    Write-Host "`n🔐 FULL ACCESS (DELEGATE) PERMISSIONS FOR: $Mailbox" -ForegroundColor Cyan
    Write-Host ("=" * 60) -ForegroundColor Cyan
    
    try {
        $Permissions = Get-MailboxPermission -Identity $Mailbox -ErrorAction Stop | Where-Object { 
            $_.User -notlike "NT AUTHORITY\SELF" -and 
            $_.User -notlike "S-1-5-*" -and
            $_.IsInherited -eq $false
        }
        
        if ($Permissions.Count -eq 0) {
            Write-Host "No Full Access permissions found." -ForegroundColor Yellow
            return
        }
        
        Write-Host "`nCurrent Full Access Permissions:" -ForegroundColor Yellow
        Write-Host ("-" * 50) -ForegroundColor Gray
        
        foreach ($Permission in $Permissions) {
            $User = $Permission.User
            $AccessRights = $Permission.AccessRights
            $Deny = if ($Permission.Deny) { " [DENY]" } else { "" }
            
            Write-Host "👤 $User$Deny" -ForegroundColor White
            Write-Host "   📋 Access Rights: $($AccessRights -join ', ')" -ForegroundColor Green
            Write-Host ""
        }
        
    } catch {
        Write-Host "❌ Error retrieving Full Access permissions: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
    
    return $true
}

# Function to display SendAs permissions
function Show-SendAsPermissions {
    param($Mailbox)
    
    Write-Host "`n📧 SEND AS PERMISSIONS FOR: $Mailbox" -ForegroundColor Cyan
    Write-Host ("=" * 60) -ForegroundColor Cyan
    
    try {
        $Permissions = Get-RecipientPermission -Identity $Mailbox -ErrorAction Stop | Where-Object { 
            $_.Trustee -notlike "NT AUTHORITY\SELF" -and
            $_.Trustee -notlike "S-1-5-*"
        }
        
        if ($Permissions.Count -eq 0) {
            Write-Host "No SendAs permissions found." -ForegroundColor Yellow
            return
        }
        
        Write-Host "`nCurrent SendAs Permissions:" -ForegroundColor Yellow
        Write-Host ("-" * 50) -ForegroundColor Gray
        
        foreach ($Permission in $Permissions) {
            $Trustee = $Permission.Trustee
            $AccessControlType = $Permission.AccessControlType
            $AccessRights = $Permission.AccessRights
            
            Write-Host "👤 $Trustee" -ForegroundColor White
            Write-Host "   📋 Access Rights: $AccessRights" -ForegroundColor Green
            Write-Host "   🔒 Access Control Type: $AccessControlType" -ForegroundColor Cyan
            Write-Host ""
        }
        
    } catch {
        Write-Host "❌ Error retrieving SendAs permissions: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
    
    return $true
}

# Function to display SendOnBehalf permissions
function Show-SendOnBehalfPermissions {
    param($Mailbox)
    
    Write-Host "`n📬 SEND ON BEHALF PERMISSIONS FOR: $Mailbox" -ForegroundColor Cyan
    Write-Host ("=" * 60) -ForegroundColor Cyan
    
    try {
        $MailboxInfo = Get-Mailbox -Identity $Mailbox -ErrorAction Stop
        $SendOnBehalf = $MailboxInfo.GrantSendOnBehalfTo
        
        if (-not $SendOnBehalf -or $SendOnBehalf.Count -eq 0) {
            Write-Host "No SendOnBehalf permissions found." -ForegroundColor Yellow
            return
        }
        
        Write-Host "`nCurrent SendOnBehalf Permissions:" -ForegroundColor Yellow
        Write-Host ("-" * 50) -ForegroundColor Gray
        
        foreach ($Delegate in $SendOnBehalf) {
            Write-Host "👤 $Delegate" -ForegroundColor White
        }
        Write-Host ""
        
    } catch {
        Write-Host "❌ Error retrieving SendOnBehalf permissions: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
    
    return $true
}

# Function to add Full Access permission
function Add-FullAccessPermission {
    param($Mailbox, $UserUPN)
    
    Write-Host "`n🔧 Adding Full Access permission..." -ForegroundColor Yellow
    Write-Host "Mailbox: $Mailbox" -ForegroundColor White
    Write-Host "User: $UserUPN" -ForegroundColor White
    
    try {
        # Check if permission already exists
        $ExistingPermission = Get-MailboxPermission -Identity $Mailbox -User $UserUPN -ErrorAction SilentlyContinue | Where-Object { 
            $_.IsInherited -eq $false -and 
            $_.AccessRights -contains "FullAccess"
        }
        
        if ($ExistingPermission) {
            Write-Host "✓ User already has Full Access permission - skipping" -ForegroundColor Green
            return $true
        }
        
        # Add new permission
        Write-Host "➕ Adding Full Access permission..." -ForegroundColor Yellow
        Add-MailboxPermission -Identity $Mailbox -User $UserUPN -AccessRights FullAccess -InheritanceType All -Confirm:$false -ErrorAction Stop
        
        Write-Host "✅ Full Access permission added successfully!" -ForegroundColor Green
        return $true
        
    } catch {
        Write-Host "❌ Error adding Full Access permission: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

# Function to remove Full Access permission
function Remove-FullAccessPermission {
    param($Mailbox, $UserUPN)
    
    Write-Host "`n🗑️  Removing Full Access permission..." -ForegroundColor Yellow
    Write-Host "Mailbox: $Mailbox" -ForegroundColor White
    Write-Host "User: $UserUPN" -ForegroundColor White
    
    try {
        # Check if permission exists
        $ExistingPermission = Get-MailboxPermission -Identity $Mailbox -User $UserUPN -ErrorAction SilentlyContinue | Where-Object { $_.IsInherited -eq $false }
        
        if (-not $ExistingPermission) {
            Write-Host "⚠️  User does not have Full Access permission to remove" -ForegroundColor Yellow
            return $false
        }
        
        Write-Host "🔄 Removing Full Access permission..." -ForegroundColor Yellow
        Remove-MailboxPermission -Identity $Mailbox -User $UserUPN -Confirm:$false -ErrorAction Stop
        
        Write-Host "✅ Full Access permission removed successfully!" -ForegroundColor Green
        return $true
        
    } catch {
        Write-Host "❌ Error removing Full Access permission: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

# Function to add SendAs permission
function Add-SendAsPermission {
    param($Mailbox, $UserUPN)
    
    Write-Host "`n🔧 Adding SendAs permission..." -ForegroundColor Yellow
    Write-Host "Mailbox: $Mailbox" -ForegroundColor White
    Write-Host "User: $UserUPN" -ForegroundColor White
    
    try {
        # Check if permission already exists
        $ExistingPermission = Get-RecipientPermission -Identity $Mailbox -Trustee $UserUPN -ErrorAction SilentlyContinue
        
        if ($ExistingPermission) {
            Write-Host "✓ User already has SendAs permission - skipping" -ForegroundColor Green
            return $true
        }
        
        # Add new permission
        Write-Host "➕ Adding SendAs permission..." -ForegroundColor Yellow
        Add-RecipientPermission -Identity $Mailbox -Trustee $UserUPN -AccessRights SendAs -Confirm:$false -ErrorAction Stop
        
        Write-Host "✅ SendAs permission added successfully!" -ForegroundColor Green
        return $true
        
    } catch {
        Write-Host "❌ Error adding SendAs permission: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

# Function to remove SendAs permission
function Remove-SendAsPermission {
    param($Mailbox, $UserUPN)
    
    Write-Host "`n🗑️  Removing SendAs permission..." -ForegroundColor Yellow
    Write-Host "Mailbox: $Mailbox" -ForegroundColor White
    Write-Host "User: $UserUPN" -ForegroundColor White
    
    try {
        # Check if permission exists
        $ExistingPermission = Get-RecipientPermission -Identity $Mailbox -Trustee $UserUPN -ErrorAction SilentlyContinue
        
        if (-not $ExistingPermission) {
            Write-Host "⚠️  User does not have SendAs permission to remove" -ForegroundColor Yellow
            return $false
        }
        
        Write-Host "🔄 Removing SendAs permission..." -ForegroundColor Yellow
        Remove-RecipientPermission -Identity $Mailbox -Trustee $UserUPN -AccessRights SendAs -Confirm:$false -ErrorAction Stop
        
        Write-Host "✅ SendAs permission removed successfully!" -ForegroundColor Green
        return $true
        
    } catch {
        Write-Host "❌ Error removing SendAs permission: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

# Function to add SendOnBehalf permission
function Add-SendOnBehalfPermission {
    param($Mailbox, $UserUPN)
    
    Write-Host "`n🔧 Adding SendOnBehalf permission..." -ForegroundColor Yellow
    Write-Host "Mailbox: $Mailbox" -ForegroundColor White
    Write-Host "User: $UserUPN" -ForegroundColor White
    
    try {
        # Get current mailbox
        $MailboxInfo = Get-Mailbox -Identity $Mailbox -ErrorAction Stop
        $CurrentDelegates = @()
        
        if ($MailboxInfo.GrantSendOnBehalfTo) {
            $CurrentDelegates = $MailboxInfo.GrantSendOnBehalfTo
        }
        
        # Check if user already has permission
        if ($CurrentDelegates -contains $UserUPN) {
            Write-Host "✓ User already has SendOnBehalf permission - skipping" -ForegroundColor Green
            return $true
        }
        
        # Add user to the list
        $NewDelegates = $CurrentDelegates + $UserUPN
        
        Write-Host "➕ Adding SendOnBehalf permission..." -ForegroundColor Yellow
        Set-Mailbox -Identity $Mailbox -GrantSendOnBehalfTo $NewDelegates -Confirm:$false -ErrorAction Stop
        
        Write-Host "✅ SendOnBehalf permission added successfully!" -ForegroundColor Green
        return $true
        
    } catch {
        Write-Host "❌ Error adding SendOnBehalf permission: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

# Function to remove SendOnBehalf permission
function Remove-SendOnBehalfPermission {
    param($Mailbox, $UserUPN)
    
    Write-Host "`n🗑️  Removing SendOnBehalf permission..." -ForegroundColor Yellow
    Write-Host "Mailbox: $Mailbox" -ForegroundColor White
    Write-Host "User: $UserUPN" -ForegroundColor White
    
    try {
        # Get current mailbox
        $MailboxInfo = Get-Mailbox -Identity $Mailbox -ErrorAction Stop
        $CurrentDelegates = @()
        
        if ($MailboxInfo.GrantSendOnBehalfTo) {
            $CurrentDelegates = $MailboxInfo.GrantSendOnBehalfTo
        }
        
        # Check if user has permission
        if ($CurrentDelegates -notcontains $UserUPN) {
            Write-Host "⚠️  User does not have SendOnBehalf permission to remove" -ForegroundColor Yellow
            return $false
        }
        
        # Remove user from the list
        $NewDelegates = $CurrentDelegates | Where-Object { $_ -ne $UserUPN }
        
        Write-Host "🔄 Removing SendOnBehalf permission..." -ForegroundColor Yellow
        Set-Mailbox -Identity $Mailbox -GrantSendOnBehalfTo $NewDelegates -Confirm:$false -ErrorAction Stop
        
        Write-Host "✅ SendOnBehalf permission removed successfully!" -ForegroundColor Green
        return $true
        
    } catch {
        Write-Host "❌ Error removing SendOnBehalf permission: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

# Function to add Full Access and SendAs permissions together
function Add-FullAccessAndSendAs {
    param($Mailbox, $UserUPN)
    
    Write-Host "`n🔧 Adding Full Access and SendAs permissions..." -ForegroundColor Yellow
    Write-Host "Mailbox: $Mailbox" -ForegroundColor White
    Write-Host "User: $UserUPN" -ForegroundColor White
    Write-Host ""
    
    Write-Host "Step 1: Adding Full Access permission..." -ForegroundColor Cyan
    $FullAccessSuccess = Add-FullAccessPermission -Mailbox $Mailbox -UserUPN $UserUPN
    
    Write-Host "`nStep 2: Adding SendAs permission..." -ForegroundColor Cyan
    $SendAsSuccess = Add-SendAsPermission -Mailbox $Mailbox -UserUPN $UserUPN
    
    if ($FullAccessSuccess -and $SendAsSuccess) {
        Write-Host "`n✅ Both Full Access and SendAs permissions added successfully!" -ForegroundColor Green
        return $true
    } elseif ($FullAccessSuccess -or $SendAsSuccess) {
        Write-Host "`n⚠️  Some permissions were added, but not all. Please check the results above." -ForegroundColor Yellow
        return $false
    } else {
        Write-Host "`n❌ Failed to add permissions." -ForegroundColor Red
        return $false
    }
}

# Function to add Full Access and SendOnBehalf permissions together
function Add-FullAccessAndSendOnBehalf {
    param($Mailbox, $UserUPN)
    
    Write-Host "`n🔧 Adding Full Access and SendOnBehalf permissions..." -ForegroundColor Yellow
    Write-Host "Mailbox: $Mailbox" -ForegroundColor White
    Write-Host "User: $UserUPN" -ForegroundColor White
    Write-Host ""
    
    Write-Host "Step 1: Adding Full Access permission..." -ForegroundColor Cyan
    $FullAccessSuccess = Add-FullAccessPermission -Mailbox $Mailbox -UserUPN $UserUPN
    
    Write-Host "`nStep 2: Adding SendOnBehalf permission..." -ForegroundColor Cyan
    $SendOnBehalfSuccess = Add-SendOnBehalfPermission -Mailbox $Mailbox -UserUPN $UserUPN
    
    if ($FullAccessSuccess -and $SendOnBehalfSuccess) {
        Write-Host "`n✅ Both Full Access and SendOnBehalf permissions added successfully!" -ForegroundColor Green
        return $true
    } elseif ($FullAccessSuccess -or $SendOnBehalfSuccess) {
        Write-Host "`n⚠️  Some permissions were added, but not all. Please check the results above." -ForegroundColor Yellow
        return $false
    } else {
        Write-Host "`n❌ Failed to add permissions." -ForegroundColor Red
        return $false
    }
}

# Main script execution
Write-Host "📧 MAILBOX PERMISSION MANAGER" -ForegroundColor Cyan
Write-Host ("=" * 50) -ForegroundColor Cyan

# Main loop - allows running multiple operations without reconnecting
do {
    # Step 1: Input mailbox
    Write-Host "`n📧 Step 1: Enter the mailbox email address" -ForegroundColor Yellow
    $Mailbox = Read-Host "Mailbox UPN (e.g., user@domain.com)"

    if ([string]::IsNullOrWhiteSpace($Mailbox)) {
        Write-Host "❌ No mailbox provided. Skipping..." -ForegroundColor Red
        $Continue = Read-Host "`nWould you like to try again? (Y/N)"
        if ($Continue -ne "Y" -and $Continue -ne "y") {
            break
        }
        continue
    }

    # Verify mailbox exists
    Write-Host "`n🔍 Verifying mailbox exists..." -ForegroundColor Yellow
    try {
        $MailboxCheck = Get-Mailbox -Identity $Mailbox -ErrorAction Stop
        Write-Host "✅ Found mailbox: $($MailboxCheck.DisplayName)" -ForegroundColor Green
    } catch {
        Write-Host "❌ Error: Could not find mailbox '$Mailbox'" -ForegroundColor Red
        Write-Host "Please verify the mailbox name and try again." -ForegroundColor Red
        $Continue = Read-Host "`nWould you like to try again? (Y/N)"
        if ($Continue -ne "Y" -and $Continue -ne "y") {
            break
        }
        continue
    }

    # Step 2: Show current permissions
    Write-Host "`n📋 Step 2: Current Mailbox Permissions" -ForegroundColor Yellow
    $FullAccessRetrieved = Show-FullAccessPermissions -Mailbox $Mailbox
    $SendAsRetrieved = Show-SendAsPermissions -Mailbox $Mailbox
    $SendOnBehalfRetrieved = Show-SendOnBehalfPermissions -Mailbox $Mailbox

    if (-not $FullAccessRetrieved -and -not $SendAsRetrieved -and -not $SendOnBehalfRetrieved) {
        Write-Host "❌ Could not retrieve mailbox permissions. Skipping..." -ForegroundColor Red
        $Continue = Read-Host "`nWould you like to try again? (Y/N)"
        if ($Continue -ne "Y" -and $Continue -ne "y") {
            break
        }
        continue
    }

    # Step 3: Ask what action to take
    Write-Host "`n❓ Step 3: What would you like to do?" -ForegroundColor Yellow
    Write-Host "ADD PERMISSIONS:" -ForegroundColor Cyan
    Write-Host "1. Add Full Access permission" -ForegroundColor White
    Write-Host "2. Add Full Access & SendAs permissions" -ForegroundColor White
    Write-Host "3. Add Full Access & SendOnBehalf permissions" -ForegroundColor White
    Write-Host "`nREMOVE PERMISSIONS:" -ForegroundColor Cyan
    Write-Host "4. Remove Full Access permission" -ForegroundColor White
    Write-Host "5. Remove SendAs permission" -ForegroundColor White
    Write-Host "6. Remove SendOnBehalf permission" -ForegroundColor White
    Write-Host "`n7. Skip this mailbox (no changes)" -ForegroundColor White

    $ValidChoice = $false
    while (-not $ValidChoice) {
        $ActionChoice = Read-Host "`nEnter your choice (1-7)"
        
        switch ($ActionChoice) {
            "1" { 
                $Action = "AddFullAccess"
                $ValidChoice = $true 
            }
            "2" { 
                $Action = "AddFullAccessAndSendAs"
                $ValidChoice = $true 
            }
            "3" { 
                $Action = "AddFullAccessAndSendOnBehalf"
                $ValidChoice = $true 
            }
            "4" { 
                $Action = "RemoveFullAccess"
                $ValidChoice = $true 
            }
            "5" { 
                $Action = "RemoveSendAs"
                $ValidChoice = $true 
            }
            "6" { 
                $Action = "RemoveSendOnBehalf"
                $ValidChoice = $true 
            }
            "7" { 
                Write-Host "`n👋 Skipping this mailbox..." -ForegroundColor Green
                $Action = "Skip"
                $ValidChoice = $true 
            }
            default { 
                Write-Host "❌ Invalid choice. Please enter 1, 2, 3, 4, 5, 6, or 7." -ForegroundColor Red
            }
        }
    }

    # Skip processing if user chose to skip
    if ($Action -eq "Skip") {
        $Continue = Read-Host "`nWould you like to manage another mailbox? (Y/N)"
        if ($Continue -ne "Y" -and $Continue -ne "y") {
            break
        }
        continue
    }

    # Step 4: Get user UPN
    if ($Action -like "Add*") {
        Write-Host "`n👤 Step 4: Enter the user who needs access" -ForegroundColor Yellow
    } else {
        Write-Host "`n👤 Step 4: Enter the user to remove access from" -ForegroundColor Yellow
    }
    $UserUPN = Read-Host "User UPN (e.g., user@domain.com)"

    if ([string]::IsNullOrWhiteSpace($UserUPN)) {
        Write-Host "❌ No user provided. Skipping..." -ForegroundColor Red
        $Continue = Read-Host "`nWould you like to try again? (Y/N)"
        if ($Continue -ne "Y" -and $Continue -ne "y") {
            break
        }
        continue
    }

    # Verify user exists
    Write-Host "`n🔍 Verifying user exists..." -ForegroundColor Yellow
    try {
        $UserCheck = Get-Mailbox -Identity $UserUPN -ErrorAction Stop
        Write-Host "✅ Found user: $($UserCheck.DisplayName)" -ForegroundColor Green
    } catch {
        Write-Host "❌ Error: Could not find user '$UserUPN'" -ForegroundColor Red
        Write-Host "Please verify the user name and try again." -ForegroundColor Red
        $Continue = Read-Host "`nWould you like to try again? (Y/N)"
        if ($Continue -ne "Y" -and $Continue -ne "y") {
            break
        }
        continue
    }

    # Step 5: Apply the permission change
    $Success = $false
    Write-Host "`n🔧 Step 5: Applying permission change..." -ForegroundColor Yellow
    
    switch ($Action) {
        "AddFullAccess" {
            $Success = Add-FullAccessPermission -Mailbox $Mailbox -UserUPN $UserUPN
        }
        "AddFullAccessAndSendAs" {
            $Success = Add-FullAccessAndSendAs -Mailbox $Mailbox -UserUPN $UserUPN
        }
        "AddFullAccessAndSendOnBehalf" {
            $Success = Add-FullAccessAndSendOnBehalf -Mailbox $Mailbox -UserUPN $UserUPN
        }
        "RemoveFullAccess" {
            $Success = Remove-FullAccessPermission -Mailbox $Mailbox -UserUPN $UserUPN
        }
        "RemoveSendAs" {
            $Success = Remove-SendAsPermission -Mailbox $Mailbox -UserUPN $UserUPN
        }
        "RemoveSendOnBehalf" {
            $Success = Remove-SendOnBehalfPermission -Mailbox $Mailbox -UserUPN $UserUPN
        }
    }

    if ($Success) {
        # Step 6: Show updated permissions
        Write-Host "`n📋 Step 6: Updated Mailbox Permissions" -ForegroundColor Yellow
        Show-FullAccessPermissions -Mailbox $Mailbox
        Show-SendAsPermissions -Mailbox $Mailbox
        Show-SendOnBehalfPermissions -Mailbox $Mailbox
        
        Write-Host "`n🎉 MAILBOX PERMISSION UPDATE COMPLETE!" -ForegroundColor Green
        Write-Host ("=" * 50) -ForegroundColor Green
        Write-Host "✅ User: $UserUPN" -ForegroundColor White
        $ActionDescription = switch ($Action) {
            "AddFullAccess" { "Added Full Access permission" }
            "AddFullAccessAndSendAs" { "Added Full Access & SendAs permissions" }
            "AddFullAccessAndSendOnBehalf" { "Added Full Access & SendOnBehalf permissions" }
            "RemoveFullAccess" { "Removed Full Access permission" }
            "RemoveSendAs" { "Removed SendAs permission" }
            "RemoveSendOnBehalf" { "Removed SendOnBehalf permission" }
            default { $Action }
        }
        Write-Host "✅ Action: $ActionDescription" -ForegroundColor White
        Write-Host "✅ Mailbox: $Mailbox" -ForegroundColor White
    } else {
        Write-Host "`n❌ Permission update failed. Please try again." -ForegroundColor Red
    }

    # Ask if user wants to run again
    Write-Host "`n" -NoNewline
    $Continue = Read-Host "Would you like to manage another mailbox? (Y/N)"
    
} while ($Continue -eq "Y" -or $Continue -eq "y")

Write-Host "`n👋 Script completed!" -ForegroundColor Cyan

#Disconnect from Exchange Online
Disconnect-ExchangeOnline