#Requires -Version 7.0

# Script: Mailbox-Distro-Report.ps1
# Purpose: Generate a comprehensive report of all mailboxes, distribution groups, mail-enabled security groups, 
#          their aliases, mailbox delegate permissions (Full Access, SendAs, SendOnBehalf), mailbox size,
#          litigation hold status, archive mailbox status, license assignment status, and account status (ACTIVE/BLOCKED)
# Author: Bobby
#
# NOTE: Only loads the minimal required Microsoft Graph submodules for faster startup. The rollup 'Microsoft.Graph' module is NOT loaded.

#region Script Configuration
$logPath = Join-Path $PSScriptRoot "logs"
$exportPath = Join-Path $PSScriptRoot "exports"

if (-not (Test-Path $logPath)) {
    New-Item -ItemType Directory -Path $logPath | Out-Null
}
if (-not (Test-Path $exportPath)) {
    New-Item -ItemType Directory -Path $exportPath | Out-Null
}

# Log file will be set dynamically per script run to match the export file date
$script:masterLogFile = $null
#endregion

#region Logging Functions
function Write-Log {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Message,
        
        [Parameter(Mandatory=$false)]
        [string]$ObjectType = "",
        
        [Parameter(Mandatory=$false)]
        [string]$ObjectName = "",
        
        [Parameter(Mandatory=$false)]
        [string]$Function = "",
        
        [Parameter(Mandatory=$false)]
        [ValidateSet("Error", "Warning", "Info")]
        [string]$ErrorType = "Error",
        
        [Parameter(Mandatory=$false)]
        [string]$ErrorDetails = ""
    )
    
    try {
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        
        $logEntry = [PSCustomObject]@{
            "Timestamp" = $timestamp
            "ObjectType" = $ObjectType
            "ObjectName" = $ObjectName
            "Function" = $Function
            "ErrorType" = $ErrorType
            "ErrorMessage" = $Message
            "ErrorDetails" = $ErrorDetails
        }
        
        # Only write log if log file has been initialized (set at start of report generation)
        if ($null -eq $script:masterLogFile) {
            Write-Host "Warning: Log file not initialized. Skipping log entry." -ForegroundColor Yellow
            return
        }
        
        # Check if file exists and has actual data rows (not just headers) to avoid duplicate headers
        $fileExists = Test-Path $script:masterLogFile -PathType Leaf
        $hasContent = $false
        
        if ($fileExists) {
            # Check if file has more than just the header row
            $fileContent = Get-Content $script:masterLogFile -ErrorAction SilentlyContinue
            if ($fileContent -and $fileContent.Count -gt 1) {
                $hasContent = $true
            }
        }
        
        if ($hasContent) {
            # File exists and has data - append data row only (skip header)
            $logEntry | ConvertTo-Csv -NoTypeInformation | Select-Object -Skip 1 | Add-Content -Path $script:masterLogFile
        }
        else {
            # File doesn't exist or only has headers - create/overwrite with header and data
            $logEntry | Export-Csv -Path $script:masterLogFile -NoTypeInformation
        }
    }
    catch {
        Write-Host "Failed to write to log file: $_" -ForegroundColor Red
    }
}
#endregion

#region Module Installation
$requiredModules = @(
    "Microsoft.Graph.Authentication",
    "Microsoft.Graph.Users",
    "Microsoft.Graph.Groups",
    "ExchangeOnlineManagement"
)

function Install-RequiredModules {
    Write-Host "Checking required modules..." -ForegroundColor Cyan
    
    $missingModules = @()
    foreach ($module in $requiredModules) {
        if (-not (Get-Module -ListAvailable -Name $module)) {
            $missingModules += $module
        }
    }
    
    if ($missingModules.Count -gt 0) {
        Write-Host "Installing missing modules: $($missingModules -join ', ')" -ForegroundColor Yellow
        foreach ($module in $missingModules) {
            try {
                Install-Module -Name $module -Scope CurrentUser -Force -ErrorAction Stop
                Write-Host "$module installed successfully" -ForegroundColor Green
            }
            catch {
                Write-Error "Failed to install $module. Error: $_"
                Write-Log -Message "Failed to install module: $module" -Function "Install-RequiredModules" -ErrorType "Error" -ErrorDetails $_.Exception.Message
                exit 1
            }
        }
    }
    else {
        Write-Host "All required modules are already installed" -ForegroundColor Green
    }
}

Install-RequiredModules

function Import-ModuleIfNeeded {
    param (
        [string]$ModuleName
    )
    
    if (-not (Get-Module -Name $ModuleName -ErrorAction SilentlyContinue)) {
        try {
            Import-Module $ModuleName -ErrorAction Stop -WarningAction SilentlyContinue
            Write-Host "✓ Imported module: $ModuleName" -ForegroundColor Green
        }
        catch {
            Write-Error "Failed to import required module '$ModuleName'. Please install it and try again."
            Write-Log -Message "Failed to import module: $ModuleName" -Function "Import-ModuleIfNeeded" -ErrorType "Error" -ErrorDetails $_.Exception.Message
            exit 1
        }
    }
}

#endregion

#region Connection Functions
function Connect-ToGraph {
    try {
        Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
        
        Import-ModuleIfNeeded "Microsoft.Graph.Authentication"
        Import-ModuleIfNeeded "Microsoft.Graph.Users"
        Import-ModuleIfNeeded "Microsoft.Graph.Groups"
        
        Write-Host "Using interactive authentication with delegated permissions..." -ForegroundColor Yellow
        Connect-MgGraph -Scopes "User.Read.All", "Group.Read.All", "Directory.Read.All", "Mail.Read" -ErrorAction Stop
        Write-Host "✓ Connected with delegated permissions (interactive login)" -ForegroundColor Green
        
        $testUser = Get-MgUser -Top 1 -ErrorAction Stop
        Write-Host "✓ Successfully retrieved user data" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Error "Failed to connect to Microsoft Graph. Error: $_"
        Write-Log -Message "Failed to connect to Microsoft Graph" -Function "Connect-ToGraph" -ErrorType "Error" -ErrorDetails $_.Exception.Message
        return $false
    }
}

function Connect-ToExchangeOnline {
    try {
        Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
        
        Import-ModuleIfNeeded "ExchangeOnlineManagement"
        
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
        
        $testMailbox = Get-Mailbox -ResultSize 1 -ErrorAction Stop
        Write-Host "✓ Successfully connected to Exchange Online" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "✗ Failed to connect to Exchange Online" -ForegroundColor Red
        Write-Host "Error: $_" -ForegroundColor Red
        Write-Log -Message "Failed to connect to Exchange Online" -Function "Connect-ToExchangeOnline" -ErrorType "Error" -ErrorDetails $_.Exception.Message
        return $false
    }
}

function Disconnect-FromGraph {
    try {
        Disconnect-MgGraph -ErrorAction SilentlyContinue
        Write-Host "Disconnected from Microsoft Graph" -ForegroundColor Yellow
    }
    catch {
        Write-Error "Error disconnecting from Microsoft Graph: $_"
        Write-Log -Message "Error disconnecting from Microsoft Graph" -Function "Disconnect-FromGraph" -ErrorType "Warning" -ErrorDetails $_.Exception.Message
    }
}

function Disconnect-FromExchangeOnline {
    try {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
        Write-Host "Disconnected from Exchange Online" -ForegroundColor Yellow
    }
    catch {
        Write-Error "Error disconnecting from Exchange Online: $_"
        Write-Log -Message "Error disconnecting from Exchange Online" -Function "Disconnect-FromExchangeOnline" -ErrorType "Warning" -ErrorDetails $_.Exception.Message
    }
}
#endregion

#region Report Generation Functions
function Get-MailboxAliases {
    param (
        [Parameter(Mandatory=$true)]
        $User
    )
    
    $aliases = @()
    
    try {
        # Get primary email for comparison
        $primaryEmail = $User.Mail
        
        # Extract all SMTP addresses from ProxyAddresses (contains both SMTP: primary and smtp: aliases)
        if ($User.ProxyAddresses) {
            foreach ($address in $User.ProxyAddresses) {
                # Extract email address from proxy format (smtp:email@domain.com or SMTP:email@domain.com)
                # Regex matches both uppercase SMTP: and lowercase smtp: prefixes
                if ($address -match "^[Ss][Mm][Tt][Pp]:(.+)$") {
                    $emailAddress = $matches[1]
                    # Exclude primary email and prevent duplicates (case-insensitive comparison)
                    if ($emailAddress -ne $primaryEmail -and 
                        $emailAddress.ToLower() -ne $primaryEmail.ToLower() -and 
                        $emailAddress -notin $aliases -and
                        $emailAddress.ToLower() -notin ($aliases | ForEach-Object { $_.ToLower() })) {
                        $aliases += $emailAddress
                    }
                }
            }
        }
        
        # Also check OtherMails property for additional email addresses
        if ($User.OtherMails) {
            foreach ($email in $User.OtherMails) {
                # Exclude primary email and prevent duplicates (case-insensitive)
                if ($email -ne $primaryEmail -and 
                    $email.ToLower() -ne $primaryEmail.ToLower() -and 
                    $email -notin $aliases -and
                    $email.ToLower() -notin ($aliases | ForEach-Object { $_.ToLower() })) {
                    $aliases += $email
                }
            }
        }
    }
    catch {
        Write-Log -Message "Failed to get aliases for mailbox" -ObjectType "Mailbox" -ObjectName $User.UserPrincipalName -Function "Get-MailboxAliases" -ErrorType "Warning" -ErrorDetails $_.Exception.Message
    }
    
    return $aliases
}

function Get-DistributionGroupAliases {
    param (
        [Parameter(Mandatory=$true)]
        $Group
    )
    
    $aliases = @()
    
    try {
        # Get primary email for comparison
        $primaryEmail = $Group.Mail
        
        # Extract all SMTP addresses from ProxyAddresses (contains both SMTP: primary and smtp: aliases)
        if ($Group.ProxyAddresses) {
            foreach ($address in $Group.ProxyAddresses) {
                # Extract email address from proxy format (smtp:email@domain.com or SMTP:email@domain.com)
                # Regex matches both uppercase SMTP: and lowercase smtp: prefixes
                if ($address -match "^[Ss][Mm][Tt][Pp]:(.+)$") {
                    $emailAddress = $matches[1]
                    # Exclude primary email and prevent duplicates (case-insensitive comparison)
                    if ($emailAddress -ne $primaryEmail -and 
                        $emailAddress.ToLower() -ne $primaryEmail.ToLower() -and 
                        $emailAddress -notin $aliases -and
                        $emailAddress.ToLower() -notin ($aliases | ForEach-Object { $_.ToLower() })) {
                        $aliases += $emailAddress
                    }
                }
            }
        }
    }
    catch {
        Write-Log -Message "Failed to get aliases for distribution group" -ObjectType "DistributionGroup" -ObjectName $Group.DisplayName -Function "Get-DistributionGroupAliases" -ErrorType "Warning" -ErrorDetails $_.Exception.Message
    }
    
    return $aliases
}

# Cache for Exchange Online connection status to avoid repeated checks
$script:ExchangeOnlineConnectionStatus = $null
$script:ExchangeOnlineLastCheck = $null

function Test-ExchangeOnlineConnection {
    # Test if Exchange Online session is still active (with caching to reduce API calls)
    $now = Get-Date
    
    # Only check if cache is older than 30 seconds or doesn't exist
    if ($null -eq $script:ExchangeOnlineLastCheck -or 
        ($now - $script:ExchangeOnlineLastCheck).TotalSeconds -gt 30) {
        try {
            $null = Get-Mailbox -ResultSize 1 -ErrorAction Stop
            $script:ExchangeOnlineConnectionStatus = $true
            $script:ExchangeOnlineLastCheck = $now
            return $true
        }
        catch {
            $script:ExchangeOnlineConnectionStatus = $false
            $script:ExchangeOnlineLastCheck = $now
            return $false
        }
    }
    
    return $script:ExchangeOnlineConnectionStatus
}

function Get-MailboxDelegates {
    param (
        [Parameter(Mandatory=$true)]
        [string]$UserPrincipalName
    )
    
    $delegates = @()
    $maxRetries = 3
    $retryCount = 0
    
    while ($retryCount -lt $maxRetries) {
        try {
            # Check if Exchange Online session is still active
            if (-not (Test-ExchangeOnlineConnection)) {
                Write-Host "Exchange Online session expired. Reconnecting..." -ForegroundColor Yellow
                if (-not (Connect-ToExchangeOnline)) {
                    Write-Log -Message "Failed to reconnect to Exchange Online" -ObjectType "Mailbox" -ObjectName $UserPrincipalName -Function "Get-MailboxDelegates" -ErrorType "Warning" -ErrorDetails "Session expired and reconnection failed"
                    return $delegates
                }
                # Reset cache after reconnection
                $script:ExchangeOnlineConnectionStatus = $true
                $script:ExchangeOnlineLastCheck = Get-Date
            }
            
            # Use Exchange Online to get mailbox permissions (Full Access)
            $permissions = Get-MailboxPermission -Identity $UserPrincipalName -ErrorAction Stop | Where-Object { 
                $_.User -notlike "NT AUTHORITY\SELF" -and 
                $_.User -notlike "S-1-5-*" -and
                $_.IsInherited -eq $false -and
                $_.AccessRights -contains "FullAccess"
            }
            
            foreach ($permission in $permissions) {
                # Try to resolve the user identity to a UPN
                try {
                    $delegateUser = Get-Mailbox -Identity $permission.User -ErrorAction SilentlyContinue
                    if ($delegateUser) {
                        $delegates += $delegateUser.UserPrincipalName
                    }
                    else {
                        # If we can't resolve, just use the raw identity
                        $delegates += $permission.User
                    }
                }
                catch {
                    # If we can't resolve, just use the raw identity
                    $delegates += $permission.User
                }
            }
            
            # Success - break out of retry loop
            break
        }
        catch {
            $retryCount++
            if ($retryCount -ge $maxRetries) {
                Write-Log -Message "Failed to get delegates for mailbox after $maxRetries attempts" -ObjectType "Mailbox" -ObjectName $UserPrincipalName -Function "Get-MailboxDelegates" -ErrorType "Warning" -ErrorDetails $_.Exception.Message
            }
            else {
                Start-Sleep -Seconds 2
            }
        }
    }
    
    return $delegates
}

function Get-MailboxSendAs {
    param (
        [Parameter(Mandatory=$true)]
        [string]$UserPrincipalName
    )
    
    $sendAsUsers = @()
    $maxRetries = 3
    $retryCount = 0
    
    while ($retryCount -lt $maxRetries) {
        try {
            # Check if Exchange Online session is still active
            if (-not (Test-ExchangeOnlineConnection)) {
                Write-Host "Exchange Online session expired. Reconnecting..." -ForegroundColor Yellow
                if (-not (Connect-ToExchangeOnline)) {
                    Write-Log -Message "Failed to reconnect to Exchange Online" -ObjectType "Mailbox" -ObjectName $UserPrincipalName -Function "Get-MailboxSendAs" -ErrorType "Warning" -ErrorDetails "Session expired and reconnection failed"
                    return $sendAsUsers
                }
                # Reset cache after reconnection
                $script:ExchangeOnlineConnectionStatus = $true
                $script:ExchangeOnlineLastCheck = Get-Date
            }
            
            # Use Exchange Online to get SendAs permissions
            $permissions = Get-RecipientPermission -Identity $UserPrincipalName -ErrorAction Stop | Where-Object { 
                $_.Trustee -notlike "NT AUTHORITY\SELF" -and
                $_.Trustee -notlike "S-1-5-*" -and
                $_.AccessRights -eq "SendAs"
            }
            
            foreach ($permission in $permissions) {
                # Try to resolve the trustee to a UPN
                try {
                    $sendAsUser = Get-Mailbox -Identity $permission.Trustee -ErrorAction SilentlyContinue
                    if ($sendAsUser) {
                        $sendAsUsers += $sendAsUser.UserPrincipalName
                    }
                    else {
                        # Try as a group
                        $sendAsGroup = Get-DistributionGroup -Identity $permission.Trustee -ErrorAction SilentlyContinue
                        if ($sendAsGroup) {
                            $sendAsUsers += $sendAsGroup.PrimarySmtpAddress
                        }
                        else {
                            # If we can't resolve, just use the raw identity
                            $sendAsUsers += $permission.Trustee
                        }
                    }
                }
                catch {
                    # If we can't resolve, just use the raw identity
                    $sendAsUsers += $permission.Trustee
                }
            }
            
            # Success - break out of retry loop
            break
        }
        catch {
            $retryCount++
            if ($retryCount -ge $maxRetries) {
                Write-Log -Message "Failed to get SendAs permissions for mailbox after $maxRetries attempts" -ObjectType "Mailbox" -ObjectName $UserPrincipalName -Function "Get-MailboxSendAs" -ErrorType "Warning" -ErrorDetails $_.Exception.Message
            }
            else {
                Start-Sleep -Seconds 2
            }
        }
    }
    
    return $sendAsUsers
}

function Get-MailboxSendOnBehalf {
    param (
        [Parameter(Mandatory=$true)]
        [string]$UserPrincipalName
    )
    
    $sendOnBehalfUsers = @()
    $maxRetries = 3
    $retryCount = 0
    
    while ($retryCount -lt $maxRetries) {
        try {
            # Check if Exchange Online session is still active
            if (-not (Test-ExchangeOnlineConnection)) {
                Write-Host "Exchange Online session expired. Reconnecting..." -ForegroundColor Yellow
                if (-not (Connect-ToExchangeOnline)) {
                    Write-Log -Message "Failed to reconnect to Exchange Online" -ObjectType "Mailbox" -ObjectName $UserPrincipalName -Function "Get-MailboxSendOnBehalf" -ErrorType "Warning" -ErrorDetails "Session expired and reconnection failed"
                    return $sendOnBehalfUsers
                }
                # Reset cache after reconnection
                $script:ExchangeOnlineConnectionStatus = $true
                $script:ExchangeOnlineLastCheck = Get-Date
            }
            
            # Use Exchange Online to get SendOnBehalf permissions
            $mailboxInfo = Get-Mailbox -Identity $UserPrincipalName -ErrorAction Stop
            
            if ($mailboxInfo -and $mailboxInfo.GrantSendOnBehalfTo) {
                foreach ($delegate in $mailboxInfo.GrantSendOnBehalfTo) {
                    # Try to resolve the delegate to a UPN
                    try {
                        $delegateUser = Get-Mailbox -Identity $delegate -ErrorAction SilentlyContinue
                        if ($delegateUser) {
                            $sendOnBehalfUsers += $delegateUser.UserPrincipalName
                        }
                        else {
                            # Try as a group
                            $delegateGroup = Get-DistributionGroup -Identity $delegate -ErrorAction SilentlyContinue
                            if ($delegateGroup) {
                                $sendOnBehalfUsers += $delegateGroup.PrimarySmtpAddress
                            }
                            else {
                                # If we can't resolve, just use the raw identity
                                $sendOnBehalfUsers += $delegate
                            }
                        }
                    }
                    catch {
                        # If we can't resolve, just use the raw identity
                        $sendOnBehalfUsers += $delegate
                    }
                }
            }
            
            # Success - break out of retry loop
            break
        }
        catch {
            $retryCount++
            if ($retryCount -ge $maxRetries) {
                Write-Log -Message "Failed to get SendOnBehalf permissions for mailbox after $maxRetries attempts" -ObjectType "Mailbox" -ObjectName $UserPrincipalName -Function "Get-MailboxSendOnBehalf" -ErrorType "Warning" -ErrorDetails $_.Exception.Message
            }
            else {
                Start-Sleep -Seconds 2
            }
        }
    }
    
    return $sendOnBehalfUsers
}

function Get-MailboxSize {
    param (
        [Parameter(Mandatory=$true)]
        [string]$UserPrincipalName
    )
    
    $mailboxSizeGB = ""
    $maxRetries = 3
    $retryCount = 0
    
    while ($retryCount -lt $maxRetries) {
        try {
            # Check if Exchange Online session is still active
            if (-not (Test-ExchangeOnlineConnection)) {
                Write-Host "Exchange Online session expired. Reconnecting..." -ForegroundColor Yellow
                if (-not (Connect-ToExchangeOnline)) {
                    Write-Log -Message "Failed to reconnect to Exchange Online" -ObjectType "Mailbox" -ObjectName $UserPrincipalName -Function "Get-MailboxSize" -ErrorType "Warning" -ErrorDetails "Session expired and reconnection failed"
                    return $mailboxSizeGB
                }
                # Reset cache after reconnection
                $script:ExchangeOnlineConnectionStatus = $true
                $script:ExchangeOnlineLastCheck = Get-Date
            }
            
            # Use Exchange Online to get mailbox statistics
            $mailboxStats = Get-MailboxStatistics -Identity $UserPrincipalName -ErrorAction Stop
            
            if ($mailboxStats -and $mailboxStats.TotalItemSize) {
                try {
                    # TotalItemSize is a ByteQuantifiedSize object
                    # Convert to bytes using the Value property's ToBytes() method
                    $sizeInBytes = $mailboxStats.TotalItemSize.Value.ToBytes()
                    
                    # Convert bytes to GB and round to 2 decimal places
                    $sizeInGB = $sizeInBytes / 1GB
                    $mailboxSizeGB = [math]::Round($sizeInGB, 2).ToString("F2")
                }
                catch {
                    # If ToBytes() fails, try parsing the string representation
                    try {
                        $sizeString = $mailboxStats.TotalItemSize.ToString()
                        # Parse string like "1.234 GB" or "1234567890 B"
                        if ($sizeString -match "([\d.]+)\s*GB") {
                            $mailboxSizeGB = [math]::Round([double]$matches[1], 2).ToString("F2")
                        }
                        elseif ($sizeString -match "([\d.]+)\s*MB") {
                            $sizeInMB = [double]$matches[1]
                            $sizeInGB = $sizeInMB / 1024
                            $mailboxSizeGB = [math]::Round($sizeInGB, 2).ToString("F2")
                        }
                        elseif ($sizeString -match "([\d.]+)\s*B") {
                            $sizeInBytes = [double]$matches[1]
                            $sizeInGB = $sizeInBytes / 1GB
                            $mailboxSizeGB = [math]::Round($sizeInGB, 2).ToString("F2")
                        }
                        else {
                            Write-Log -Message "Unable to parse mailbox size format: $sizeString" -ObjectType "Mailbox" -ObjectName $UserPrincipalName -Function "Get-MailboxSize" -ErrorType "Warning" -ErrorDetails "Unknown size format"
                        }
                    }
                    catch {
                        Write-Log -Message "Failed to parse mailbox size" -ObjectType "Mailbox" -ObjectName $UserPrincipalName -Function "Get-MailboxSize" -ErrorType "Warning" -ErrorDetails $_.Exception.Message
                    }
                }
            }
            elseif ($mailboxStats -and $null -eq $mailboxStats.TotalItemSize) {
                # Mailbox exists but has no items
                $mailboxSizeGB = "0.00"
            }
            
            # Success - break out of retry loop
            break
        }
        catch {
            $retryCount++
            if ($retryCount -ge $maxRetries) {
                Write-Log -Message "Failed to get mailbox size after $maxRetries attempts" -ObjectType "Mailbox" -ObjectName $UserPrincipalName -Function "Get-MailboxSize" -ErrorType "Warning" -ErrorDetails $_.Exception.Message
            }
            else {
                Start-Sleep -Seconds 2
            }
        }
    }
    
    return $mailboxSizeGB
}

function Get-MailboxLitigationHold {
    param (
        [Parameter(Mandatory=$true)]
        [string]$UserPrincipalName
    )
    
    $litigationHold = ""
    $maxRetries = 3
    $retryCount = 0
    
    while ($retryCount -lt $maxRetries) {
        try {
            # Check if Exchange Online session is still active
            if (-not (Test-ExchangeOnlineConnection)) {
                Write-Host "Exchange Online session expired. Reconnecting..." -ForegroundColor Yellow
                if (-not (Connect-ToExchangeOnline)) {
                    Write-Log -Message "Failed to reconnect to Exchange Online" -ObjectType "Mailbox" -ObjectName $UserPrincipalName -Function "Get-MailboxLitigationHold" -ErrorType "Warning" -ErrorDetails "Session expired and reconnection failed"
                    return $litigationHold
                }
                # Reset cache after reconnection
                $script:ExchangeOnlineConnectionStatus = $true
                $script:ExchangeOnlineLastCheck = Get-Date
            }
            
            # Use Exchange Online to get mailbox properties
            $mailbox = Get-Mailbox -Identity $UserPrincipalName -ErrorAction Stop
            
            if ($mailbox) {
                # LitigationHoldEnabled is a boolean property
                if ($mailbox.LitigationHoldEnabled -eq $true) {
                    $litigationHold = "True"
                }
                else {
                    $litigationHold = "False"
                }
            }
            
            # Success - break out of retry loop
            break
        }
        catch {
            $retryCount++
            if ($retryCount -ge $maxRetries) {
                Write-Log -Message "Failed to get litigation hold status after $maxRetries attempts" -ObjectType "Mailbox" -ObjectName $UserPrincipalName -Function "Get-MailboxLitigationHold" -ErrorType "Warning" -ErrorDetails $_.Exception.Message
            }
            else {
                Start-Sleep -Seconds 2
            }
        }
    }
    
    return $litigationHold
}

function Get-MailboxArchiveStatus {
    param (
        [Parameter(Mandatory=$true)]
        [string]$UserPrincipalName
    )
    
    $archiveEnabled = ""
    $maxRetries = 3
    $retryCount = 0
    
    while ($retryCount -lt $maxRetries) {
        try {
            # Check if Exchange Online session is still active
            if (-not (Test-ExchangeOnlineConnection)) {
                Write-Host "Exchange Online session expired. Reconnecting..." -ForegroundColor Yellow
                if (-not (Connect-ToExchangeOnline)) {
                    Write-Log -Message "Failed to reconnect to Exchange Online" -ObjectType "Mailbox" -ObjectName $UserPrincipalName -Function "Get-MailboxArchiveStatus" -ErrorType "Warning" -ErrorDetails "Session expired and reconnection failed"
                    return $archiveEnabled
                }
                # Reset cache after reconnection
                $script:ExchangeOnlineConnectionStatus = $true
                $script:ExchangeOnlineLastCheck = Get-Date
            }
            
            # Use Exchange Online to get mailbox properties
            $mailbox = Get-Mailbox -Identity $UserPrincipalName -ErrorAction Stop
            
            if ($mailbox) {
                # ArchiveStatus property indicates if archive is enabled
                # ArchiveDatabase property exists if archive is enabled
                if ($mailbox.ArchiveStatus -eq "Active" -or ($null -ne $mailbox.ArchiveDatabase -and $mailbox.ArchiveDatabase -ne "")) {
                    $archiveEnabled = "True"
                }
                else {
                    $archiveEnabled = "False"
                }
            }
            
            # Success - break out of retry loop
            break
        }
        catch {
            $retryCount++
            if ($retryCount -ge $maxRetries) {
                Write-Log -Message "Failed to get archive mailbox status after $maxRetries attempts" -ObjectType "Mailbox" -ObjectName $UserPrincipalName -Function "Get-MailboxArchiveStatus" -ErrorType "Warning" -ErrorDetails $_.Exception.Message
            }
            else {
                Start-Sleep -Seconds 2
            }
        }
    }
    
    return $archiveEnabled
}

function Get-UserLicenseStatus {
    param (
        [Parameter(Mandatory=$true)]
        [string]$UserPrincipalName
    )
    
    $hasLicense = ""
    $maxRetries = 3
    $retryCount = 0
    
    while ($retryCount -lt $maxRetries) {
        try {
            # Use Microsoft Graph to get user license information
            $user = Get-MgUser -UserId $UserPrincipalName -Property Id, AssignedLicenses -ErrorAction Stop
            
            if ($user) {
                # Check if user has any assigned licenses
                if ($user.AssignedLicenses -and $user.AssignedLicenses.Count -gt 0) {
                    $hasLicense = "True"
                }
                else {
                    $hasLicense = "False"
                }
            }
            
            # Success - break out of retry loop
            break
        }
        catch {
            $retryCount++
            if ($retryCount -ge $maxRetries) {
                Write-Log -Message "Failed to get license status after $maxRetries attempts" -ObjectType "Mailbox" -ObjectName $UserPrincipalName -Function "Get-UserLicenseStatus" -ErrorType "Warning" -ErrorDetails $_.Exception.Message
            }
            else {
                Start-Sleep -Seconds 2
            }
        }
    }
    
    return $hasLicense
}

function Get-UserAccountStatus {
    param (
        [Parameter(Mandatory=$true)]
        [string]$UserPrincipalName
    )
    
    $accountStatus = ""
    $maxRetries = 3
    $retryCount = 0
    
    while ($retryCount -lt $maxRetries) {
        try {
            # Use Microsoft Graph to get user account status
            $user = Get-MgUser -UserId $UserPrincipalName -Property Id, AccountEnabled -ErrorAction Stop
            
            if ($user) {
                # Check if account is enabled (ACTIVE) or disabled (BLOCKED)
                if ($user.AccountEnabled -eq $true) {
                    $accountStatus = "ACTIVE"
                }
                else {
                    $accountStatus = "BLOCKED"
                }
            }
            
            # Success - break out of retry loop
            break
        }
        catch {
            $retryCount++
            if ($retryCount -ge $maxRetries) {
                Write-Log -Message "Failed to get account status after $maxRetries attempts" -ObjectType "Mailbox" -ObjectName $UserPrincipalName -Function "Get-UserAccountStatus" -ErrorType "Warning" -ErrorDetails $_.Exception.Message
            }
            else {
                Start-Sleep -Seconds 2
            }
        }
    }
    
    return $accountStatus
}

function Export-MailboxDistroReport {
    param (
        [Parameter(Mandatory=$false)]
        [string]$ExportPath = $exportPath
    )
    
    $reportData = @()
    $friendlyDate = Get-Date -Format "MM-dd-yyyy"
    $exportFile = Join-Path $ExportPath "Mailbox-Distro-Report-$friendlyDate.csv"
    
    # Set log file to match export file date (unique per script run)
    $script:masterLogFile = Join-Path $logPath "ReportErrors-$friendlyDate.csv"
    
    $errorCount = 0
    $startTime = Get-Date
    
    Write-Host "`n==== Starting Report Generation ====" -ForegroundColor Cyan
    Write-Host "Report will be saved to: $exportFile" -ForegroundColor Yellow
    Write-Host "Log file: $script:masterLogFile" -ForegroundColor Gray
    Write-Host "Start time: $($startTime.ToString('yyyy-MM-dd HH:mm:ss'))" -ForegroundColor Gray
    
    # Get all mailboxes (user mailboxes and shared mailboxes, excluding guest accounts)
    Write-Host "`nRetrieving mailboxes..." -ForegroundColor Cyan
    Write-Host "Note: This may take a while for large environments. The script will process ALL mailboxes including shared mailboxes (excluding guest accounts)." -ForegroundColor Yellow
    
    $allMailboxObjects = @()
    $processedMailboxEmails = @{}
    
    try {
        # First, get user mailboxes from Graph API
        Write-Host "Retrieving user mailboxes from Microsoft Graph..." -ForegroundColor Cyan
        $allUsers = Get-MgUser -All -Property Id, DisplayName, Mail, UserPrincipalName, ProxyAddresses, OtherMails, UserType, AccountEnabled -ErrorAction Stop
        $totalUsers = ($allUsers | Measure-Object).Count
        
        # Filter to only mailboxes (users with Mail property) and exclude guest accounts
        $userMailboxes = $allUsers | Where-Object { 
            $_.Mail -and 
            ($_.UserType -ne "Guest" -or $null -eq $_.UserType)
        }
        $userMailboxCount = ($userMailboxes | Measure-Object).Count
        $guestCount = ($allUsers | Where-Object { $_.UserType -eq "Guest" } | Measure-Object).Count
        
        Write-Host "Found $totalUsers total users ($guestCount guest accounts excluded), $userMailboxCount user mailboxes" -ForegroundColor Green
        
        # Add user mailboxes to collection
        foreach ($mailbox in $userMailboxes) {
            $allMailboxObjects += [PSCustomObject]@{
                DisplayName = $mailbox.DisplayName
                PrimaryEmail = $mailbox.Mail
                UserPrincipalName = $mailbox.UserPrincipalName
                ProxyAddresses = $mailbox.ProxyAddresses
                OtherMails = $mailbox.OtherMails
                MailboxType = "UserMailbox"
                AccountEnabled = $mailbox.AccountEnabled
            }
            $processedMailboxEmails[$mailbox.Mail] = $true
        }
        
        # Get shared mailboxes from Exchange Online (if connected)
        $sharedMailboxesAdded = 0
        $sharedMailboxesSkipped = 0
        if (Test-ExchangeOnlineConnection) {
            Write-Host "Retrieving shared mailboxes from Exchange Online..." -ForegroundColor Cyan
            try {
                # Get shared mailboxes - EmailAddresses property contains all addresses (SMTP, SIP, X400, etc.)
                # Note: Get-Mailbox returns EmailAddresses by default; ProxyAddresses may also be available
                $sharedMailboxes = Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited -ErrorAction Stop
                $sharedCount = ($sharedMailboxes | Measure-Object).Count
                Write-Host "Found $sharedCount shared mailboxes" -ForegroundColor Green
                
                foreach ($sharedMb in $sharedMailboxes) {
                    # Skip if we already processed this mailbox (avoid duplicates)
                    if (-not $processedMailboxEmails.ContainsKey($sharedMb.PrimarySmtpAddress)) {
                        # Extract ALL SMTP addresses from EmailAddresses property (includes both SMTP: and smtp: prefixes)
                        $proxyAddresses = @()
                        if ($sharedMb.EmailAddresses) {
                            foreach ($addr in $sharedMb.EmailAddresses) {
                                # Match both "SMTP:" (primary) and "smtp:" (aliases) - primary will be filtered out later
                                if ($addr -match "^[Ss][Mm][Tt][Pp]:") {
                                    $proxyAddresses += $addr
                                }
                            }
                        }
                        
                        # Also check ProxyAddresses property if available
                        if ($sharedMb.ProxyAddresses) {
                            foreach ($addr in $sharedMb.ProxyAddresses) {
                                if ($addr -match "^[Ss][Mm][Tt][Pp]:") {
                                    if ($proxyAddresses -notcontains $addr) {
                                        $proxyAddresses += $addr
                                    }
                                }
                            }
                        }
                        
                        $allMailboxObjects += [PSCustomObject]@{
                            DisplayName = $sharedMb.DisplayName
                            PrimaryEmail = $sharedMb.PrimarySmtpAddress
                            UserPrincipalName = $sharedMb.UserPrincipalName
                            ProxyAddresses = $proxyAddresses
                            OtherMails = @()
                            MailboxType = "SharedMailbox"
                        }
                        $processedMailboxEmails[$sharedMb.PrimarySmtpAddress] = $true
                        $sharedMailboxesAdded++
                    }
                    else {
                        $sharedMailboxesSkipped++
                    }
                }
                
                if ($sharedMailboxesSkipped -gt 0) {
                    Write-Host "  Added $sharedMailboxesAdded shared mailboxes, skipped $sharedMailboxesSkipped duplicates" -ForegroundColor Gray
                }
            }
            catch {
                Write-Host "Warning: Could not retrieve shared mailboxes from Exchange Online: $_" -ForegroundColor Yellow
                Write-Log -Message "Could not retrieve shared mailboxes" -Function "Export-MailboxDistroReport" -ErrorType "Warning" -ErrorDetails $_.Exception.Message
            }
        }
        else {
            Write-Host "Exchange Online not connected - skipping shared mailbox retrieval" -ForegroundColor Yellow
        }
        
        $mailboxCount = ($allMailboxObjects | Measure-Object).Count
        $userMailboxCount = ($allMailboxObjects | Where-Object { $_.MailboxType -eq "UserMailbox" } | Measure-Object).Count
        $sharedMailboxCount = ($allMailboxObjects | Where-Object { $_.MailboxType -eq "SharedMailbox" } | Measure-Object).Count
        Write-Host "Total mailboxes to process: $mailboxCount ($userMailboxCount user mailboxes + $sharedMailboxCount shared mailboxes)" -ForegroundColor Green
        
        $currentMailbox = 0
        foreach ($mailboxObj in $allMailboxObjects) {
            $currentMailbox++
            $percentComplete = [math]::Round(($currentMailbox / $mailboxCount) * 100, 1)
            $elapsed = (Get-Date) - $startTime
            $avgTimePerMailbox = $elapsed.TotalSeconds / $currentMailbox
            $remainingMailboxes = $mailboxCount - $currentMailbox
            $estimatedRemaining = [TimeSpan]::FromSeconds($avgTimePerMailbox * $remainingMailboxes)
            
            if ($currentMailbox % 50 -eq 0 -or $currentMailbox -eq 1) {
                $mailboxTypeLabel = if ($mailboxObj.MailboxType -eq "SharedMailbox") { "shared mailbox" } else { "mailbox" }
                Write-Host "Processing $mailboxTypeLabel $currentMailbox of $mailboxCount ($percentComplete%) - Elapsed: $($elapsed.ToString('hh\:mm\:ss')) - Est. remaining: $($estimatedRemaining.ToString('hh\:mm\:ss'))" -ForegroundColor Yellow
            }
            
            # Initialize variables
            $mailboxSizeGB = ""
            $litigationHold = ""
            $archiveEnabled = ""
            $hasLicense = ""
            $accountStatus = ""
            
            try {
                # Extract all SMTP aliases (exclude primary email address)
                $aliases = @()
                $primaryEmail = $mailboxObj.PrimaryEmail
                
                # Process ProxyAddresses property (contains all SMTP addresses)
                if ($mailboxObj.ProxyAddresses) {
                    foreach ($address in $mailboxObj.ProxyAddresses) {
                        # Extract email address from proxy format (smtp:email@domain.com or SMTP:email@domain.com)
                        # Regex matches both uppercase SMTP: and lowercase smtp: prefixes
                        if ($address -match "^[Ss][Mm][Tt][Pp]:(.+)$") {
                            $emailAddress = $matches[1]
                            # Exclude primary email and prevent duplicates (case-insensitive comparison)
                            if ($emailAddress -ne $primaryEmail -and 
                                $emailAddress.ToLower() -ne $primaryEmail.ToLower() -and 
                                $emailAddress -notin $aliases -and
                                $emailAddress.ToLower() -notin ($aliases | ForEach-Object { $_.ToLower() })) {
                                $aliases += $emailAddress
                            }
                        }
                    }
                }
                
                # Also check OtherMails property for additional email addresses
                if ($mailboxObj.OtherMails) {
                    foreach ($email in $mailboxObj.OtherMails) {
                        # Exclude primary email and prevent duplicates (case-insensitive)
                        if ($email -ne $primaryEmail -and 
                            $email.ToLower() -ne $primaryEmail.ToLower() -and 
                            $email -notin $aliases -and
                            $email.ToLower() -notin ($aliases | ForEach-Object { $_.ToLower() })) {
                            $aliases += $email
                        }
                    }
                }
                
                # Get delegate permissions (Full Access, SendAs, SendOnBehalf) for mailboxes
                $delegates = Get-MailboxDelegates -UserPrincipalName $mailboxObj.UserPrincipalName
                $sendAs = Get-MailboxSendAs -UserPrincipalName $mailboxObj.UserPrincipalName
                $sendOnBehalf = Get-MailboxSendOnBehalf -UserPrincipalName $mailboxObj.UserPrincipalName
                
                # Get mailbox size
                $mailboxSizeGB = Get-MailboxSize -UserPrincipalName $mailboxObj.UserPrincipalName
                
                # Get litigation hold status
                $litigationHold = Get-MailboxLitigationHold -UserPrincipalName $mailboxObj.UserPrincipalName
                
                # Get archive mailbox status
                $archiveEnabled = Get-MailboxArchiveStatus -UserPrincipalName $mailboxObj.UserPrincipalName
                
                # Get license assignment status
                $hasLicense = Get-UserLicenseStatus -UserPrincipalName $mailboxObj.UserPrincipalName
                
                # Get account status (ACTIVE/BLOCKED) - use cached value if available, otherwise fetch
                if ($null -ne $mailboxObj.AccountEnabled) {
                    # Use cached AccountEnabled value from initial user retrieval
                    if ($mailboxObj.AccountEnabled -eq $true) {
                        $accountStatus = "ACTIVE"
                    }
                    else {
                        $accountStatus = "BLOCKED"
                    }
                }
                else {
                    # Fallback: fetch account status if not cached (for shared mailboxes or edge cases)
                    $accountStatus = Get-UserAccountStatus -UserPrincipalName $mailboxObj.UserPrincipalName
                }
            }
            catch {
                $errorCount++
                Write-Host "Warning: Failed to process mailbox $($mailboxObj.UserPrincipalName): $_" -ForegroundColor Yellow
                Write-Log -Message "Failed to process mailbox" -ObjectType "Mailbox" -ObjectName $mailboxObj.UserPrincipalName -Function "Export-MailboxDistroReport" -ErrorType "Warning" -ErrorDetails $_.Exception.Message
                # Continue processing other mailboxes even if one fails
                $aliases = @()
                $delegates = @()
                $sendAs = @()
                $sendOnBehalf = @()
                $mailboxSizeGB = ""
                $litigationHold = ""
                $archiveEnabled = ""
                $hasLicense = ""
                $accountStatus = ""
            }
            
            # Small delay every 100 mailboxes to avoid API rate limiting
            if ($currentMailbox % 100 -eq 0) {
                Start-Sleep -Milliseconds 500
            }
            
            $aliasesString = if ($aliases.Count -gt 0) { ($aliases -join "; ") } else { "" }
            $delegatesString = if ($delegates.Count -gt 0) { ($delegates -join "; ") } else { "" }
            $sendAsString = if ($sendAs.Count -gt 0) { ($sendAs -join "; ") } else { "" }
            $sendOnBehalfString = if ($sendOnBehalf.Count -gt 0) { ($sendOnBehalf -join "; ") } else { "" }
            
            # Set mailbox type label for CSV report
            $mailboxTypeLabel = if ($mailboxObj.MailboxType -eq "SharedMailbox") { "Shared Mailbox" } else { "Mailbox" }
            
            $reportData += [PSCustomObject]@{
                "AccountStatus" = $accountStatus
                "Type" = $mailboxTypeLabel
                "DisplayName" = $mailboxObj.DisplayName
                "PrimaryEmail" = $mailboxObj.PrimaryEmail
                "UserPrincipalName" = $mailboxObj.UserPrincipalName
                "Aliases" = $aliasesString
                "Delegates" = $delegatesString
                "SendAs" = $sendAsString
                "SendOnBehalf" = $sendOnBehalfString
                "MailboxSizeGB" = $mailboxSizeGB
                "LitigationHold" = $litigationHold
                "ArchiveEnabled" = $archiveEnabled
                "LicenseAssigned" = $hasLicense
            }
        }
        $mailboxElapsed = (Get-Date) - $startTime
        Write-Host "✓ Completed processing $mailboxCount mailboxes in $($mailboxElapsed.ToString('hh\:mm\:ss'))" -ForegroundColor Green
        if ($errorCount -gt 0) {
            Write-Host "  ⚠ $errorCount mailboxes had errors (check logs for details)" -ForegroundColor Yellow
        }
    }
    catch {
        Write-Host "✗ Failed to retrieve mailboxes: $_" -ForegroundColor Red
        Write-Log -Message "Failed to retrieve mailboxes" -Function "Export-MailboxDistroReport" -ErrorType "Error" -ErrorDetails $_.Exception.Message
    }
    
    # Get all mail-enabled groups (distribution groups and mail-enabled security groups, excluding Microsoft 365 groups)
    Write-Host "`nRetrieving mail-enabled groups..." -ForegroundColor Cyan
    Write-Host "Note: This may take a while for large environments. The script will process ALL groups." -ForegroundColor Yellow
    try {
        # Use -All parameter to retrieve all groups (handles pagination automatically)
        $allMailGroups = Get-MgGroup -All -Filter "mailEnabled eq true" -Property Id, DisplayName, Mail, ProxyAddresses, GroupTypes, SecurityEnabled -ErrorAction Stop
        
        # Filter out Microsoft 365 groups (Unified groups) - keep only distribution groups and mail-enabled security groups
        $mailGroups = $allMailGroups | Where-Object { 
            $_.GroupTypes -notcontains "Unified" -and 
            $null -ne $_.Mail
        }
        
        $distroCount = ($mailGroups | Measure-Object).Count
        Write-Host "Found $distroCount mail-enabled groups (distribution groups and mail-enabled security groups)" -ForegroundColor Green
        
        $groupStartTime = Get-Date
        $currentDistro = 0
        foreach ($distro in $mailGroups) {
            $currentDistro++
            $percentComplete = [math]::Round(($currentDistro / $distroCount) * 100, 1)
            $elapsed = (Get-Date) - $groupStartTime
            
            if ($currentDistro % 50 -eq 0 -or $currentDistro -eq 1) {
                Write-Host "Processing mail-enabled group $currentDistro of $distroCount ($percentComplete%) - Elapsed: $($elapsed.ToString('hh\:mm\:ss'))" -ForegroundColor Yellow
            }
            
            # Skip groups without a mail property
            if (-not $distro.Mail) {
                continue
            }
            
            # Set group type label for CSV report
            $groupType = if ($distro.SecurityEnabled -eq $true) {
                "Mail-Enabled Security Group"
            } else {
                "Distribution Group"
            }
            
            $aliases = Get-DistributionGroupAliases -Group $distro
            $aliasesString = if ($aliases.Count -gt 0) { ($aliases -join "; ") } else { "" }
            
            $reportData += [PSCustomObject]@{
                "AccountStatus" = ""
                "Type" = $groupType
                "DisplayName" = $distro.DisplayName
                "PrimaryEmail" = $distro.Mail
                "UserPrincipalName" = ""
                "Aliases" = $aliasesString
                "Delegates" = ""
                "SendAs" = ""
                "SendOnBehalf" = ""
                "MailboxSizeGB" = ""
                "LitigationHold" = ""
                "ArchiveEnabled" = ""
                "LicenseAssigned" = ""
            }
        }
        $groupElapsed = (Get-Date) - $groupStartTime
        Write-Host "✓ Completed processing $distroCount mail-enabled groups in $($groupElapsed.ToString('hh\:mm\:ss'))" -ForegroundColor Green
    }
    catch {
        Write-Host "✗ Failed to retrieve distribution groups: $_" -ForegroundColor Red
        Write-Log -Message "Failed to retrieve distribution groups" -Function "Export-MailboxDistroReport" -ErrorType "Error" -ErrorDetails $_.Exception.Message
    }
    
    # Export to CSV
    Write-Host "`nExporting report to CSV..." -ForegroundColor Cyan
    try {
        if ($reportData.Count -gt 0) {
            $reportData | Export-Csv -Path $exportFile -NoTypeInformation -Encoding UTF8
            $totalElapsed = (Get-Date) - $startTime
            Write-Host "✓ Successfully exported $($reportData.Count) objects to: $exportFile" -ForegroundColor Green
            Write-Host "`nReport Summary:" -ForegroundColor Cyan
            Write-Host "  - User Mailboxes: $(($reportData | Where-Object { $_.Type -eq 'Mailbox' }).Count)" -ForegroundColor White
            Write-Host "  - Shared Mailboxes: $(($reportData | Where-Object { $_.Type -eq 'Shared Mailbox' }).Count)" -ForegroundColor White
            Write-Host "  - Distribution Groups: $(($reportData | Where-Object { $_.Type -eq 'Distribution Group' }).Count)" -ForegroundColor White
            Write-Host "  - Mail-Enabled Security Groups: $(($reportData | Where-Object { $_.Type -eq 'Mail-Enabled Security Group' }).Count)" -ForegroundColor White
            Write-Host "  - Total Objects: $($reportData.Count)" -ForegroundColor White
            Write-Host "  - Total Processing Time: $($totalElapsed.ToString('hh\:mm\:ss'))" -ForegroundColor White
            if ($errorCount -gt 0) {
                Write-Host "  - Errors Encountered: $errorCount (check logs for details)" -ForegroundColor Yellow
            }
            return $exportFile
        }
        else {
            Write-Host "✗ No data to export" -ForegroundColor Red
            return $null
        }
    }
    catch {
        Write-Host "✗ Failed to export report: $_" -ForegroundColor Red
        Write-Log -Message "Failed to export report" -Function "Export-MailboxDistroReport" -ErrorType "Error" -ErrorDetails $_.Exception.Message
        return $null
    }
}
#endregion

#region Main Function
function Main {
    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host "  Mailbox & Distribution Group Report  " -ForegroundColor Cyan
    Write-Host "========================================`n" -ForegroundColor Cyan
    
    Write-Host "Establishing required connections..." -ForegroundColor Cyan
    
    $connectedToGraph = Connect-ToGraph
    if (!$connectedToGraph) {
        Write-Host "✗ Failed to connect to Microsoft Graph. Cannot proceed with report generation." -ForegroundColor Red
        return
    }
    
    $connectedToExchange = Connect-ToExchangeOnline
    if (!$connectedToExchange) {
        Write-Host "⚠ Warning: Failed to connect to Exchange Online. Delegate information, mailbox size, litigation hold, and archive status will not be available." -ForegroundColor Yellow
        Write-Host "Continuing with report generation (aliases, license information, and account status from Graph API will be available)..." -ForegroundColor Yellow
    }
    
    Write-Host "✓ All required connections established successfully" -ForegroundColor Green
    
    # Generate the report
    $reportFile = Export-MailboxDistroReport
    
    if ($reportFile) {
        Write-Host "`n========================================" -ForegroundColor Green
        Write-Host "  Report Generation Complete!           " -ForegroundColor Green
        Write-Host "========================================`n" -ForegroundColor Green
        Write-Host "Report saved to: $reportFile" -ForegroundColor Yellow
        
        # Ask if user wants to open the file
        $openFile = Read-Host "`nWould you like to open the report file? (Y/N)"
        if ($openFile -eq "Y" -or $openFile -eq "y") {
            try {
                Invoke-Item $reportFile
            }
            catch {
                Write-Host "Could not automatically open the file. Please open it manually at: $reportFile" -ForegroundColor Yellow
            }
        }
    }
    else {
        Write-Host "`n✗ Report generation failed. Please check the logs for details." -ForegroundColor Red
    }
}
#endregion

# --- Script Entry Point ---
# Temporarily suppress warnings to hide benign module loading messages
$OldWarningPreference = $WarningPreference
$WarningPreference = 'SilentlyContinue'

try {
    Main
}
finally {
    # This block always runs, ensuring disconnection and restoration of settings
    Write-Host "`nDisconnecting from services and restoring settings..." -ForegroundColor Yellow
    Disconnect-FromGraph
    Disconnect-FromExchangeOnline
    $WarningPreference = $OldWarningPreference
}

