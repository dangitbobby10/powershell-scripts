# PowerShell Script for Interactive Calendar Permission Management
# This script allows you to view and modify calendar permissions for any mailbox

#Connect to Exchange Online
Connect-ExchangeOnline

# Function to display calendar permissions
function Show-CalendarPermissions {
    param($Mailbox)
    
    Write-Host "`n📅 CALENDAR PERMISSIONS FOR: $Mailbox" -ForegroundColor Cyan
    Write-Host ("=" * 60) -ForegroundColor Cyan
    
    try {
        $Permissions = Get-MailboxFolderPermission -Identity "$Mailbox`:\calendar" -ErrorAction Stop
        
        if ($Permissions.Count -eq 0) {
            Write-Host "No calendar permissions found." -ForegroundColor Yellow
            return
        }
        
        Write-Host "`nCurrent Calendar Permissions:" -ForegroundColor Yellow
        Write-Host ("-" * 50) -ForegroundColor Gray
        
        foreach ($Permission in $Permissions) {
            $User = $Permission.User
            $AccessRights = $Permission.AccessRights
            $IsInherited = $Permission.IsInherited
            $SharingPermissionFlags = $Permission.SharingPermissionFlags
            
            # Skip default/system permissions
            if ($User -like "Default" -or $User -like "Anonymous" -or $IsInherited -eq $true) {
                continue
            }
            
            $Status = if ($IsInherited) { "(Inherited)" } else { "" }
            $IsDelegate = if ($SharingPermissionFlags -like "*Delegate*") { " [DELEGATE]" } else { "" }
            
            Write-Host "👤 $User$IsDelegate" -ForegroundColor White
            Write-Host "   📋 Access Level: $AccessRights $Status" -ForegroundColor Green
            
            # Show delegate-specific flags
            if ($SharingPermissionFlags) {
                $Flags = @()
                if ($SharingPermissionFlags -like "*CanViewPrivateItems*") {
                    $Flags += "View Private Events"
                }
                if ($SharingPermissionFlags -like "*CanManageCategories*") {
                    $Flags += "Manage Categories"
                }
                if ($Flags.Count -gt 0) {
                    Write-Host "   🔐 Delegate Options: $($Flags -join ', ')" -ForegroundColor Cyan
                }
            }
            Write-Host ""
        }
        
    } catch {
        Write-Host "❌ Error retrieving calendar permissions: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
    
    return $true
}

# Function to display delegates separately
function Show-Delegates {
    param($Mailbox)
    
    Write-Host "`n👥 DELEGATES FOR: $Mailbox" -ForegroundColor Cyan
    Write-Host ("=" * 60) -ForegroundColor Cyan
    
    try {
        $Permissions = Get-MailboxFolderPermission -Identity "$Mailbox`:\calendar" -ErrorAction Stop
        
        $Delegates = $Permissions | Where-Object { 
            $_.SharingPermissionFlags -like "*Delegate*" -and 
            $_.User -notlike "Default" -and 
            $_.User -notlike "Anonymous" -and 
            $_.IsInherited -eq $false 
        }
        
        if ($Delegates.Count -eq 0) {
            Write-Host "No delegates found." -ForegroundColor Yellow
            return
        }
        
        Write-Host "`nCurrent Delegates:" -ForegroundColor Yellow
        Write-Host ("-" * 50) -ForegroundColor Gray
        
        foreach ($Delegate in $Delegates) {
            $User = $Delegate.User
            $AccessRights = $Delegate.AccessRights
            $SharingPermissionFlags = $Delegate.SharingPermissionFlags
            
            Write-Host "👤 $User" -ForegroundColor White
            Write-Host "   📋 Access Level: $AccessRights" -ForegroundColor Green
            
            $Options = @()
            if ($SharingPermissionFlags -like "*CanViewPrivateItems*") {
                $Options += "✓ View Private Events"
            } else {
                $Options += "✗ View Private Events"
            }
            if ($SharingPermissionFlags -like "*CanManageCategories*") {
                $Options += "✓ Manage Categories"
            } else {
                $Options += "✗ Manage Categories"
            }
            
            Write-Host "   🔐 Options: $($Options -join ' | ')" -ForegroundColor Cyan
            Write-Host ""
        }
        
    } catch {
        Write-Host "❌ Error retrieving delegates: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
    
    return $true
}

# Function to display default calendar sharing permission
function Show-DefaultPermission {
    param($Mailbox)
    
    Write-Host "`n🏢 DEFAULT SHARING PERMISSION (Inside your organization)" -ForegroundColor Cyan
    Write-Host ("=" * 60) -ForegroundColor Cyan
    
    try {
        $DefaultPermission = Get-MailboxFolderPermission -Identity "$Mailbox`:\calendar" -User "Default" -ErrorAction Stop
        
        if ($DefaultPermission) {
            $AccessRights = $DefaultPermission.AccessRights
            Write-Host "`nCurrent Default Permission:" -ForegroundColor Yellow
            Write-Host ("-" * 50) -ForegroundColor Gray
            Write-Host "👥 People in my organization" -ForegroundColor White
            Write-Host "   📋 Access Level: $AccessRights" -ForegroundColor Green
            Write-Host ""
            return $AccessRights
        } else {
            Write-Host "No default permission found (using system default)." -ForegroundColor Yellow
            return $null
        }
        
    } catch {
        Write-Host "❌ Error retrieving default permission: $($_.Exception.Message)" -ForegroundColor Red
        return $null
    }
}

# Function to set default calendar sharing permission
function Set-DefaultPermission {
    param($Mailbox, $AccessLevel)
    
    Write-Host "`n🔧 Setting default calendar sharing permission..." -ForegroundColor Yellow
    Write-Host "Mailbox: $Mailbox" -ForegroundColor White
    Write-Host "Default Access Level: $AccessLevel" -ForegroundColor White
    Write-Host "This affects: People in my organization" -ForegroundColor White
    
    try {
        # Check if default permission already exists
        $ExistingPermission = Get-MailboxFolderPermission -Identity "$Mailbox`:\calendar" -User "Default" -ErrorAction SilentlyContinue
        
        if ($ExistingPermission) {
            Write-Host "⚠️  Default permission already exists: $($ExistingPermission.AccessRights)" -ForegroundColor Yellow
            Write-Host "🔄 Updating default permission..." -ForegroundColor Yellow
            
            # Update existing permission
            Set-MailboxFolderPermission -Identity "$Mailbox`:\calendar" -User "Default" -AccessRights $AccessLevel -Confirm:$false
        } else {
            Write-Host "➕ Adding default permission..." -ForegroundColor Yellow
            # Add new default permission
            Add-MailboxFolderPermission -Identity "$Mailbox`:\calendar" -User "Default" -AccessRights $AccessLevel -Confirm:$false
        }
        
        Write-Host "✅ Default permission set successfully!" -ForegroundColor Green
        return $true
        
    } catch {
        Write-Host "❌ Error setting default permission: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

# Function to display available permission levels
function Show-PermissionLevels {
    Write-Host "`n📋 AVAILABLE CALENDAR PERMISSION LEVELS:" -ForegroundColor Yellow
    Write-Host "-" * 50 -ForegroundColor Gray
    
    $PermissionLevels = @{
        "1" = @{ Name = "Owner"; Description = "Full access - read, create, modify, delete all items and folders" }
        "2" = @{ Name = "PublishingEditor"; Description = "Read, create, modify, delete items/subfolders" }
        "3" = @{ Name = "Editor"; Description = "Read, create, modify, delete items" }
        "4" = @{ Name = "PublishingAuthor"; Description = "Read, create items/subfolders. Modify/delete own items only" }
        "5" = @{ Name = "Author"; Description = "Create and read items; edit/delete own items" }
        "6" = @{ Name = "NonEditingAuthor"; Description = "Read access and create items. Delete own items only" }
        "7" = @{ Name = "Reviewer"; Description = "Read only access" }
        "8" = @{ Name = "Contributor"; Description = "Create items and folders only" }
        "9" = @{ Name = "AvailabilityOnly"; Description = "Read free/busy information only --- DEFAULT SETTING" }
        "10" = @{ Name = "LimitedDetails"; Description = "View subject and location only --- TITLES ONLY, NO DETAILS" }
    }
    
    foreach ($Key in $PermissionLevels.Keys | Sort-Object { [int]$_ }) {
        $Level = $PermissionLevels[$Key]
        Write-Host "  $Key. $($Level.Name)" -ForegroundColor White
        Write-Host "     $($Level.Description)" -ForegroundColor Gray
        Write-Host ""
    }
}

# Function to add/update calendar permission
function Set-CalendarPermission {
    param($Mailbox, $UserUPN, $AccessLevel, [string[]]$SharingPermissionFlags = @())
    
    Write-Host "`n🔧 Setting calendar permission..." -ForegroundColor Yellow
    Write-Host "Mailbox: $Mailbox" -ForegroundColor White
    Write-Host "User: $UserUPN" -ForegroundColor White
    Write-Host "Access Level: $AccessLevel" -ForegroundColor White
    if ($SharingPermissionFlags.Count -gt 0) {
        Write-Host "Delegate Flags: $($SharingPermissionFlags -join ', ')" -ForegroundColor White
    }
    
    try {
        # First, check if user already has permissions
        $ExistingPermission = Get-MailboxFolderPermission -Identity "$Mailbox`:\calendar" -User $UserUPN -ErrorAction SilentlyContinue
        
        if ($ExistingPermission) {
            Write-Host "⚠️  User already has calendar access: $($ExistingPermission.AccessRights)" -ForegroundColor Yellow
            Write-Host "🔄 Removing existing permission..." -ForegroundColor Yellow
            
            # Remove existing permission
            Remove-MailboxFolderPermission -Identity "$Mailbox`:\calendar" -User $UserUPN -Confirm:$false
            Write-Host "✅ Existing permission removed" -ForegroundColor Green
        }
        
        # Add new permission
        Write-Host "➕ Adding new permission..." -ForegroundColor Yellow
        
        if ($SharingPermissionFlags.Count -gt 0) {
            Add-MailboxFolderPermission -Identity "$Mailbox`:\calendar" -User $UserUPN -AccessRights $AccessLevel -SharingPermissionFlags $SharingPermissionFlags -Confirm:$false
        } else {
            Add-MailboxFolderPermission -Identity "$Mailbox`:\calendar" -User $UserUPN -AccessRights $AccessLevel -Confirm:$false
        }
        
        Write-Host "✅ Permission set successfully!" -ForegroundColor Green
        return $true
        
    } catch {
        Write-Host "❌ Error setting permission: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

# Function to add/update delegate permission
function Set-DelegatePermission {
    param($Mailbox, $UserUPN, $AccessLevel, $ViewPrivateEvents, $ManageCategories)
    
    Write-Host "`n🔧 Setting delegate permission..." -ForegroundColor Yellow
    Write-Host "Mailbox: $Mailbox" -ForegroundColor White
    Write-Host "Delegate: $UserUPN" -ForegroundColor White
    Write-Host "Access Level: $AccessLevel" -ForegroundColor White
    Write-Host "View Private Events: $ViewPrivateEvents" -ForegroundColor White
    Write-Host "Manage Categories: $ManageCategories" -ForegroundColor White
    
    try {
        # Build SharingPermissionFlags array
        $Flags = @("Delegate")
        if ($ViewPrivateEvents) {
            $Flags += "CanViewPrivateItems"
        }
        if ($ManageCategories) {
            $Flags += "CanManageCategories"
        }
        
        # Use Set-CalendarPermission with delegate flags
        return Set-CalendarPermission -Mailbox $Mailbox -UserUPN $UserUPN -AccessLevel $AccessLevel -SharingPermissionFlags $Flags
        
    } catch {
        Write-Host "❌ Error setting delegate permission: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

# Function to remove calendar permission
function Remove-CalendarPermission {
    param($Mailbox, $UserUPN)
    
    Write-Host "`n🗑️  Removing calendar permission..." -ForegroundColor Yellow
    Write-Host "Mailbox: $Mailbox" -ForegroundColor White
    Write-Host "User: $UserUPN" -ForegroundColor White
    
    try {
        # Check if user has permissions
        $ExistingPermission = Get-MailboxFolderPermission -Identity "$Mailbox`:\calendar" -User $UserUPN -ErrorAction SilentlyContinue
        
        if (-not $ExistingPermission) {
            Write-Host "⚠️  User does not have calendar access to remove" -ForegroundColor Yellow
            return $false
        }
        
        Write-Host "📋 Current access level: $($ExistingPermission.AccessRights)" -ForegroundColor White
        Write-Host "🔄 Removing calendar permission..." -ForegroundColor Yellow
        
        # Remove permission
        Remove-MailboxFolderPermission -Identity "$Mailbox`:\calendar" -User $UserUPN -Confirm:$false
        Write-Host "✅ Calendar permission removed successfully!" -ForegroundColor Green
        return $true
        
    } catch {
        Write-Host "❌ Error removing permission: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

# Main script execution
Write-Host "🗓️  CALENDAR PERMISSION MANAGER" -ForegroundColor Cyan
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

    # Step 2: Show current calendar permissions and delegates
    Write-Host "`n📅 Step 2: Current Calendar Permissions" -ForegroundColor Yellow
    $PermissionsRetrieved = Show-CalendarPermissions -Mailbox $Mailbox

    if (-not $PermissionsRetrieved) {
        Write-Host "❌ Could not retrieve calendar permissions. Skipping..." -ForegroundColor Red
        $Continue = Read-Host "`nWould you like to try again? (Y/N)"
        if ($Continue -ne "Y" -and $Continue -ne "y") {
            break
        }
        continue
    }

    Write-Host "`n👥 Step 2b: Current Delegates" -ForegroundColor Yellow
    Show-Delegates -Mailbox $Mailbox

    Write-Host "`n🏢 Step 2c: Default Sharing Permission" -ForegroundColor Yellow
    Show-DefaultPermission -Mailbox $Mailbox

    # Step 3: Ask what action to take
    Write-Host "`n❓ Step 3: What would you like to do?" -ForegroundColor Yellow
    Write-Host "1. Add/Update calendar access for someone" -ForegroundColor White
    Write-Host "2. Remove calendar access for someone" -ForegroundColor White
    Write-Host "3. Add/Update delegate access" -ForegroundColor White
    Write-Host "4. Remove delegate access" -ForegroundColor White
    Write-Host "5. Set default sharing permission (Inside your organization)" -ForegroundColor White
    Write-Host "6. Skip this mailbox (no changes)" -ForegroundColor White

    $ValidChoice = $false
    while (-not $ValidChoice) {
        $ActionChoice = Read-Host "`nEnter your choice (1-6)"
        
        switch ($ActionChoice) {
            "1" { 
                $Action = "Add"
                $ValidChoice = $true 
            }
            "2" { 
                $Action = "Remove"
                $ValidChoice = $true 
            }
            "3" { 
                $Action = "AddDelegate"
                $ValidChoice = $true 
            }
            "4" { 
                $Action = "RemoveDelegate"
                $ValidChoice = $true 
            }
            "5" { 
                $Action = "SetDefault"
                $ValidChoice = $true 
            }
            "6" { 
                Write-Host "`n👋 Skipping this mailbox..." -ForegroundColor Green
                $Action = "Skip"
                $ValidChoice = $true 
            }
            default { 
                Write-Host "❌ Invalid choice. Please enter 1, 2, 3, 4, 5, or 6." -ForegroundColor Red
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

    # Step 4: Get user UPN (skip for SetDefault action)
    if ($Action -ne "SetDefault") {
        if ($Action -eq "Add" -or $Action -eq "AddDelegate") {
            if ($Action -eq "AddDelegate") {
                Write-Host "`n👤 Step 4: Enter the delegate who needs access" -ForegroundColor Yellow
            } else {
                Write-Host "`n👤 Step 4: Enter the user who needs access" -ForegroundColor Yellow
            }
            $UserUPN = Read-Host "User UPN (e.g., newuser@domain.com)"
        } else {
            if ($Action -eq "RemoveDelegate") {
                Write-Host "`n👤 Step 4: Enter the delegate to remove access from" -ForegroundColor Yellow
            } else {
                Write-Host "`n👤 Step 4: Enter the user to remove access from" -ForegroundColor Yellow
            }
            $UserUPN = Read-Host "User UPN (e.g., user@domain.com)"
        }

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
    }

    if ($Action -eq "Add") {
        # Step 5: Select permission level (only for Add action)
        Write-Host "`n🔐 Step 5: Select permission level" -ForegroundColor Yellow
        Show-PermissionLevels

        $ValidChoice = $false
        while (-not $ValidChoice) {
            $Choice = Read-Host "`nEnter choice (1-10)"
            
            $PermissionMap = @{
                "1" = "Owner"
                "2" = "PublishingEditor"
                "3" = "Editor"
                "4" = "PublishingAuthor"
                "5" = "Author"
                "6" = "NonEditingAuthor"
                "7" = "Reviewer"
                "8" = "Contributor"
                "9" = "AvailabilityOnly"
                "10" = "LimitedDetails"
            }
            
            if ($PermissionMap.ContainsKey($Choice)) {
                $AccessLevel = $PermissionMap[$Choice]
                $ValidChoice = $true
            } else {
                Write-Host "❌ Invalid choice. Please enter a number between 1-10." -ForegroundColor Red
            }
        }

        # Apply permission
        Write-Host "`n🔧 Step 6: Applying permission..." -ForegroundColor Yellow
        $Success = Set-CalendarPermission -Mailbox $Mailbox -UserUPN $UserUPN -AccessLevel $AccessLevel

        if ($Success) {
            # Step 7: Show updated permissions
            Write-Host "`n📅 Step 7: Updated Calendar Permissions" -ForegroundColor Yellow
            Show-CalendarPermissions -Mailbox $Mailbox
            
            Write-Host "`n🎉 CALENDAR PERMISSION UPDATE COMPLETE!" -ForegroundColor Green
            Write-Host ("=" * 50) -ForegroundColor Green
            Write-Host "✅ User: $UserUPN" -ForegroundColor White
            Write-Host "✅ Access Level: $AccessLevel" -ForegroundColor White
            Write-Host "✅ Mailbox: $Mailbox" -ForegroundColor White
        } else {
            Write-Host "`n❌ Permission update failed. Please try again." -ForegroundColor Red
        }
    } elseif ($Action -eq "AddDelegate") {
        # Step 5: Select permission level for delegate
        Write-Host "`n🔐 Step 5: Select permission level for delegate" -ForegroundColor Yellow
        Write-Host "Note: Delegates typically need Editor or higher access" -ForegroundColor Gray
        Show-PermissionLevels

        $ValidChoice = $false
        while (-not $ValidChoice) {
            $Choice = Read-Host "`nEnter choice (1-10)"
            
            $PermissionMap = @{
                "1" = "Owner"
                "2" = "PublishingEditor"
                "3" = "Editor"
                "4" = "PublishingAuthor"
                "5" = "Author"
                "6" = "NonEditingAuthor"
                "7" = "Reviewer"
                "8" = "Contributor"
                "9" = "AvailabilityOnly"
                "10" = "LimitedDetails"
            }
            
            if ($PermissionMap.ContainsKey($Choice)) {
                $AccessLevel = $PermissionMap[$Choice]
                $ValidChoice = $true
            } else {
                Write-Host "❌ Invalid choice. Please enter a number between 1-10." -ForegroundColor Red
            }
        }

        # Step 6: Delegate options
        Write-Host "`n🔐 Step 6: Configure delegate options" -ForegroundColor Yellow
        Write-Host "Delegates can view, create, modify and delete items." -ForegroundColor Gray
        Write-Host "They can also create meeting requests and respond to invitations on your behalf." -ForegroundColor Gray
        Write-Host ""
        
        $ViewPrivateChoice = Read-Host "Let delegate view private events? (Y/N)"
        $ViewPrivateEvents = ($ViewPrivateChoice -eq "Y" -or $ViewPrivateChoice -eq "y")
        
        $ManageCategoriesChoice = Read-Host "Let delegate manage categories? (Y/N)"
        $ManageCategories = ($ManageCategoriesChoice -eq "Y" -or $ManageCategoriesChoice -eq "y")

        # Apply delegate permission
        Write-Host "`n🔧 Step 7: Applying delegate permission..." -ForegroundColor Yellow
        $Success = Set-DelegatePermission -Mailbox $Mailbox -UserUPN $UserUPN -AccessLevel $AccessLevel -ViewPrivateEvents $ViewPrivateEvents -ManageCategories $ManageCategories

        if ($Success) {
            # Step 8: Show updated permissions and delegates
            Write-Host "`n📅 Step 8: Updated Calendar Permissions" -ForegroundColor Yellow
            Show-CalendarPermissions -Mailbox $Mailbox
            Show-Delegates -Mailbox $Mailbox
            
            Write-Host "`n🎉 DELEGATE PERMISSION UPDATE COMPLETE!" -ForegroundColor Green
            Write-Host ("=" * 50) -ForegroundColor Green
            Write-Host "✅ Delegate: $UserUPN" -ForegroundColor White
            Write-Host "✅ Access Level: $AccessLevel" -ForegroundColor White
            Write-Host "✅ View Private Events: $ViewPrivateEvents" -ForegroundColor White
            Write-Host "✅ Manage Categories: $ManageCategories" -ForegroundColor White
            Write-Host "✅ Mailbox: $Mailbox" -ForegroundColor White
            Write-Host "`n⚠️  Note: 'Send invitations and responses to' setting must be configured in Outlook." -ForegroundColor Yellow
        } else {
            Write-Host "`n❌ Delegate permission update failed. Please try again." -ForegroundColor Red
        }
    } elseif ($Action -eq "RemoveDelegate") {
        # Remove delegate permission (same as regular remove, but shows delegate info)
        Write-Host "`n🔧 Step 5: Removing delegate permission..." -ForegroundColor Yellow
        $Success = Remove-CalendarPermission -Mailbox $Mailbox -UserUPN $UserUPN

        if ($Success) {
            # Step 6: Show updated permissions
            Write-Host "`n📅 Step 6: Updated Calendar Permissions" -ForegroundColor Yellow
            Show-CalendarPermissions -Mailbox $Mailbox
            Show-Delegates -Mailbox $Mailbox
            
            Write-Host "`n🎉 DELEGATE PERMISSION REMOVAL COMPLETE!" -ForegroundColor Green
            Write-Host ("=" * 50) -ForegroundColor Green
            Write-Host "✅ Delegate: $UserUPN" -ForegroundColor White
            Write-Host "✅ Action: Delegate access removed" -ForegroundColor White
            Write-Host "✅ Mailbox: $Mailbox" -ForegroundColor White
        } else {
            Write-Host "`n❌ Delegate permission removal failed. Please try again." -ForegroundColor Red
        }
    } elseif ($Action -eq "SetDefault") {
        # Step 4: Select default permission level
        Write-Host "`n🔐 Step 4: Select default sharing permission level" -ForegroundColor Yellow
        Write-Host "This sets the permission for 'People in my organization'" -ForegroundColor Gray
        Write-Host "Common choices: AvailabilityOnly (default) or LimitedDetails" -ForegroundColor Gray
        Show-PermissionLevels

        $ValidChoice = $false
        while (-not $ValidChoice) {
            $Choice = Read-Host "`nEnter choice (1-10)"
            
            $PermissionMap = @{
                "1" = "Owner"
                "2" = "PublishingEditor"
                "3" = "Editor"
                "4" = "PublishingAuthor"
                "5" = "Author"
                "6" = "NonEditingAuthor"
                "7" = "Reviewer"
                "8" = "Contributor"
                "9" = "AvailabilityOnly"
                "10" = "LimitedDetails"
            }
            
            if ($PermissionMap.ContainsKey($Choice)) {
                $AccessLevel = $PermissionMap[$Choice]
                $ValidChoice = $true
            } else {
                Write-Host "❌ Invalid choice. Please enter a number between 1-10." -ForegroundColor Red
            }
        }

        # Apply default permission
        Write-Host "`n🔧 Step 5: Applying default permission..." -ForegroundColor Yellow
        $Success = Set-DefaultPermission -Mailbox $Mailbox -AccessLevel $AccessLevel

        if ($Success) {
            # Step 6: Show updated default permission
            Write-Host "`n🏢 Step 6: Updated Default Sharing Permission" -ForegroundColor Yellow
            Show-DefaultPermission -Mailbox $Mailbox
            
            Write-Host "`n🎉 DEFAULT PERMISSION UPDATE COMPLETE!" -ForegroundColor Green
            Write-Host ("=" * 50) -ForegroundColor Green
            Write-Host "✅ Default Access Level: $AccessLevel" -ForegroundColor White
            Write-Host "✅ Applies to: People in my organization" -ForegroundColor White
            Write-Host "✅ Mailbox: $Mailbox" -ForegroundColor White
        } else {
            Write-Host "`n❌ Default permission update failed. Please try again." -ForegroundColor Red
        }
    } else {
        # Remove permission
        Write-Host "`n🔧 Step 5: Removing permission..." -ForegroundColor Yellow
        $Success = Remove-CalendarPermission -Mailbox $Mailbox -UserUPN $UserUPN

        if ($Success) {
            # Step 6: Show updated permissions
            Write-Host "`n📅 Step 6: Updated Calendar Permissions" -ForegroundColor Yellow
            Show-CalendarPermissions -Mailbox $Mailbox
            
            Write-Host "`n🎉 CALENDAR PERMISSION REMOVAL COMPLETE!" -ForegroundColor Green
            Write-Host ("=" * 50) -ForegroundColor Green
            Write-Host "✅ User: $UserUPN" -ForegroundColor White
            Write-Host "✅ Action: Calendar access removed" -ForegroundColor White
            Write-Host "✅ Mailbox: $Mailbox" -ForegroundColor White
        } else {
            Write-Host "`n❌ Permission removal failed. Please try again." -ForegroundColor Red
        }
    }

    # Ask if user wants to run again
    Write-Host "`n" -NoNewline
    $Continue = Read-Host "Would you like to manage another mailbox? (Y/N)"
    
} while ($Continue -eq "Y" -or $Continue -eq "y")

Write-Host "`n👋 Script completed!" -ForegroundColor Cyan

#Disconnect from Exchange Online
Disconnect-ExchangeOnline