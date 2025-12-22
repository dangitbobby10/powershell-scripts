# Remove script parameters
#Requires -Version 7.0
# #Requires -Modules @{ModuleName="Microsoft.Graph"; ModuleVersion="1.0.0"}

# Script: EntraTermination1.ps1 (Optimized)
# Purpose: Automate user termination process in Microsoft Entra ID
# Author: Bobby
#
# NOTE: This version only loads the minimal required Microsoft Graph submodules for faster startup. The rollup 'Microsoft.Graph' module is NOT loaded.

#region Module Installation
$requiredModules = @(
    "Microsoft.Graph.Authentication",
    "Microsoft.Graph.Identity.DirectoryManagement",
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
                exit 1
            }
        }
    }
    else {
        Write-Host "All required modules are already installed" -ForegroundColor Green
    }
}

# Only install modules if they're missing
Install-RequiredModules

# Lazy load modules as needed
function Import-ModuleIfNeeded {
    param (
        [string]$ModuleName
    )
    
    # Check if the module is already loaded to avoid re-importing
    if (-not (Get-Module -Name $ModuleName -ErrorAction SilentlyContinue)) {
        try {
            Import-Module $ModuleName -ErrorAction Stop -WarningAction SilentlyContinue
            Write-Host "✓ Imported module: $ModuleName" -ForegroundColor Green
        }
        catch {
            Write-Error "Failed to import required module '$ModuleName'. Please install it and try again."
            exit 1
        }
    }
}

#endregion

#region Script Configuration
# Add configuration variables here
$logPath = Join-Path $PSScriptRoot "logs"
$exportPath = Join-Path $PSScriptRoot "exports"

# Create necessary directories if they don't exist
if (-not (Test-Path $logPath)) {
    New-Item -ItemType Directory -Path $logPath | Out-Null
}
if (-not (Test-Path $exportPath)) {
    New-Item -ItemType Directory -Path $exportPath | Out-Null
}
#endregion

#region Connection Functions
function Connect-ToGraph {
    try {
        Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
        
        # Import required modules
        Import-ModuleIfNeeded "Microsoft.Graph.Authentication"
        Import-ModuleIfNeeded "Microsoft.Graph.Users"
        Import-ModuleIfNeeded "Microsoft.Graph.Identity.DirectoryManagement"
        Import-ModuleIfNeeded "Microsoft.Graph.Groups"
        
        # Using interactive authentication with delegated permissions
        Write-Host "Using interactive authentication with delegated permissions..." -ForegroundColor Yellow
        Connect-MgGraph -Scopes "User.ReadWrite.All", "Group.ReadWrite.All", "Directory.ReadWrite.All", "Mail.ReadWrite", "MailboxSettings.ReadWrite", "User.ManageIdentities.All", "User.Read.All", "User.ReadWrite.All" -ErrorAction Stop
        Write-Host "✓ Connected with delegated permissions (interactive login)" -ForegroundColor Green
        
        # Test the connection
        $testUser = Get-MgUser -Top 1 -ErrorAction Stop
        Write-Host "✓ Successfully retrieved user data" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Error "Failed to connect to Microsoft Graph. Error: $_"
        return $false
    }
}

function Connect-ToExchangeOnline {
    try {
        Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
        
        # Import required module
        Import-ModuleIfNeeded "ExchangeOnlineManagement"
        
        # Exchange Online requires interactive authentication
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
        
        # Test the connection
        $testMailbox = Get-Mailbox -ResultSize 1 -ErrorAction Stop
        Write-Host "✓ Successfully connected to Exchange Online" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "✗ Failed to connect to Exchange Online" -ForegroundColor Red
        Write-Host "Error: $_" -ForegroundColor Red
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
    }
}

function Disconnect-FromExchangeOnline {
    try {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
        Write-Host "Disconnected from Exchange Online" -ForegroundColor Yellow
    }
    catch {
        Write-Error "Error disconnecting from Exchange Online: $_"
    }
}

function Revoke-UserSessions {
    param (
        [Parameter(Mandatory=$true)]
        [string]$UserPrincipalName
    )
    
    try {
        Write-Host "Revoking active sessions for $UserPrincipalName..." -ForegroundColor Cyan
        
        # Get the user ID
        $user = Get-MgUser -UserId $UserPrincipalName -ErrorAction Stop
        $userId = $user.Id
        
        # Use the Microsoft Graph PowerShell cmdlet for revoking sessions
        Revoke-MgUserSignInSession -UserId $userId -ErrorAction Stop
        
        Write-Host "✓ Successfully revoked all active sessions for $UserPrincipalName" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "✗ Failed to revoke sessions for $UserPrincipalName" -ForegroundColor Red
        Write-Host "Error: $_" -ForegroundColor Red
        return $false
    }
}
#endregion

#region Verification Functions
function Test-ModuleInstallation {
    Write-Host "`nVerifying module installation..." -ForegroundColor Cyan
    $missingModules = @()
    
    foreach ($module in $requiredModules) {
        $installed = Get-Module -ListAvailable -Name $module
        if ($installed) {
            Write-Host "✓ $module is installed (Version: $($installed.Version))" -ForegroundColor Green
        }
        else {
            Write-Host "✗ $module is NOT installed" -ForegroundColor Red
            $missingModules += $module
        }
    }
    
    if ($missingModules.Count -gt 0) {
        Write-Host "`nMissing modules:" -ForegroundColor Yellow
        $missingModules | ForEach-Object { Write-Host "- $_" -ForegroundColor Yellow }
        return $false
    }
    return $true
}
#endregion

#region Functions
# Note: Password reset functionality has been removed due to AzureAD module compatibility issues on Mac

function Block-UserSignIn {
    param (
        [Parameter(Mandatory=$true)]
        [string]$UserPrincipalName
    )
    
    try {
        Write-Host "Blocking sign-in for $UserPrincipalName..." -ForegroundColor Cyan
        
        # First, get the user ID and check their current state
        $user = Get-MgUser -UserId $UserPrincipalName -ErrorAction Stop
        $userId = $user.Id
        
        Write-Host "Found user with ID: $userId" -ForegroundColor Cyan
        Write-Host "Current account state: $($user.AccountEnabled)" -ForegroundColor Cyan
        
        # Robust check for already disabled (boolean or string)
        if (($user.AccountEnabled -eq $false) -or ($user.AccountEnabled -eq "False")) {
            Write-Host "Account is already disabled." -ForegroundColor Yellow
            return $true
        }
        
        # Get user's directory roles
        $userRoles = Get-MgUserMemberOf -UserId $userId -ErrorAction Stop
        $adminRoles = $userRoles | Where-Object { $_.AdditionalProperties["@odata.type"] -eq "#microsoft.graph.directoryRole" }
        
        if ($adminRoles) {
            Write-Host "User has admin roles. Attempting to remove admin roles first..." -ForegroundColor Yellow
            
            # First try to remove admin roles
            foreach ($role in $adminRoles) {
                try {
                    Remove-MgDirectoryRoleMemberByRef -DirectoryRoleId $role.Id -DirectoryObjectId $userId -ErrorAction Stop
                    Write-Host "Removed admin role: $($role.DisplayName)" -ForegroundColor Green
                }
                catch {
                    Write-Host "Warning: Could not remove admin role: $($role.DisplayName)" -ForegroundColor Yellow
                    Write-Host "Error: $_" -ForegroundColor Yellow
                }
            }
        }
        
        # Block sign-in using Update-MgUser
        $params = @{
            AccountEnabled = $false
        }
        
        Write-Host "Attempting to disable account..." -ForegroundColor Cyan
        Update-MgUser -UserId $userId -BodyParameter $params -ErrorAction Stop
        
        return $true
    }
    catch {
        Write-Host "✗ Failed to block sign-in for $UserPrincipalName" -ForegroundColor Red
        Write-Host "Error: $_" -ForegroundColor Red
        Write-Host "Error Type: $($_.Exception.GetType().FullName)" -ForegroundColor Red
        Write-Host "Error Message: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

function Hide-UserFromGAL {
    param (
        [Parameter(Mandatory=$true)]
        [string]$UserPrincipalName
    )
    
    try {
        Write-Host "Hiding $UserPrincipalName from Global Address List using Exchange Online..." -ForegroundColor Cyan
        
        # Hide from GAL using Exchange Online command
        Set-Mailbox -Identity $UserPrincipalName -HiddenFromAddressListsEnabled $true -ErrorAction Stop
        
        Write-Host "✓ Successfully hidden $UserPrincipalName from Global Address List" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "✗ Failed to hide $UserPrincipalName from Global Address List" -ForegroundColor Red
        Write-Host "Error: $_" -ForegroundColor Red
        return $false
    }
}

function Convert-ToSharedMailbox {
    param (
        [Parameter(Mandatory=$true)]
        [string]$UserPrincipalName
    )
    
    try {
        Write-Host "Converting $UserPrincipalName to shared mailbox..." -ForegroundColor Cyan
        
        # Convert to shared mailbox using Exchange Online PowerShell
        Set-Mailbox -Identity $UserPrincipalName -Type Shared -ErrorAction Stop
        
        Write-Host "✓ Successfully converted $UserPrincipalName to a shared mailbox" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "✗ Failed to convert $UserPrincipalName to a shared mailbox" -ForegroundColor Red
        Write-Host "Error: $_" -ForegroundColor Red
        return $false
    }
}

function Rename-UserDisplayName {
    param (
        [Parameter(Mandatory=$true)]
        [string]$UserPrincipalName
    )
    
    try {
        Write-Host "Renaming display name for $UserPrincipalName..." -ForegroundColor Cyan
        
        # Get the user and current display name
        $user = Get-MgUser -UserId $UserPrincipalName -ErrorAction Stop
        $userId = $user.Id
        $currentDisplayName = $user.DisplayName
        
        # Check if the name already contains "- Email Archive" to avoid duplication
        if ($currentDisplayName -like "*- Email Archive*") {
            Write-Host "Display name already contains '- Email Archive'. No changes needed." -ForegroundColor Yellow
            return $true
        }
        
        # Format the new display name to "[Current Name] - Email Archive"
        $newDisplayName = "$currentDisplayName - Email Archive"
        
        # Update the display name
        Update-MgUser -UserId $userId -DisplayName $newDisplayName -ErrorAction Stop
        
        Write-Host "✓ Successfully renamed display name to: $newDisplayName" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "✗ Failed to rename display name for $UserPrincipalName" -ForegroundColor Red
        Write-Host "Error: $_" -ForegroundColor Red
        return $false
    }
}

function Export-UserMemberships {
    param (
        [Parameter(Mandatory=$true)]
        [string]$UserPrincipalName,
        
        [Parameter(Mandatory=$false)]
        [string]$ExportPath = (Join-Path $PSScriptRoot "exports")
    )
    
    try {
        Write-Host "Exporting group memberships for $UserPrincipalName..." -ForegroundColor Cyan
        
        # Ensure the export directory exists
        if (-not (Test-Path $ExportPath)) {
            New-Item -ItemType Directory -Path $ExportPath | Out-Null
        }
        
        # Get user details
        $user = Get-MgUser -UserId $UserPrincipalName -ErrorAction Stop
        $userId = $user.Id
        $upnForFilename = $UserPrincipalName
        $friendlyDate = Get-Date -Format "MM-dd-yyyy"
        $exportFile = Join-Path $ExportPath ("${upnForFilename}_offboarding-$friendlyDate.csv")
        
        Write-Host "Found user with ID: $userId" -ForegroundColor Cyan
        
        # Get group memberships
        Write-Host "Getting group memberships..." -ForegroundColor Cyan
        $groupMemberships = @()
        try {
            # Get user's group memberships
            $groups = Get-MgUserMemberOf -UserId $userId -All -ErrorAction Stop
            
            # Process each group
            foreach ($group in $groups) {
                if ($group.AdditionalProperties["@odata.type"] -eq "#microsoft.graph.group") {
                    # Get the group ID
                    $groupId = $group.Id
                    
                    # Get group details
                    $groupDetails = Get-MgGroup -GroupId $groupId -ErrorAction SilentlyContinue
                    
                    if ($groupDetails) {
                        # Determine group type
                        $groupType = "Security"
                        if ($groupDetails.GroupTypes -contains "Unified") {
                            $groupType = "Microsoft 365"
                        }
                        elseif ($groupDetails.MailEnabled -eq $true) {
                            $groupType = "Mail-Enabled Security"
                            if ($groupDetails.SecurityEnabled -ne $true) {
                                $groupType = "Distribution"
                            }
                        }
                        
                        $groupMemberships += [PSCustomObject]@{
                            "Group Type" = $groupType
                            "Group Name" = $groupDetails.DisplayName
                        }
                    }
                }
            }
        }
        catch {
            Write-Host "Warning: Could not retrieve group memberships: $_" -ForegroundColor Yellow
        }
        
        # Export to CSV
        if ($groupMemberships.Count -gt 0) {
            $groupMemberships | Export-Csv -Path $exportFile -NoTypeInformation
            Write-Host "✓ Successfully exported $($groupMemberships.Count) group memberships to: $exportFile" -ForegroundColor Green
        }
        else {
            Write-Host "No group memberships found for $UserPrincipalName" -ForegroundColor Yellow
            # Create empty file with headers
            [PSCustomObject]@{
                "Group Type" = ""
                "Group Name" = ""
            } | Export-Csv -Path $exportFile -NoTypeInformation
        }
        
        return $exportFile
    }
    catch {
        Write-Host "✗ Failed to export group memberships for $UserPrincipalName" -ForegroundColor Red
        Write-Host "Error: $_" -ForegroundColor Red
        return $null
    }
}

#region License Mapping Cache
$script:LicenseMappingCache = @{}

function Get-LicenseFriendlyName {
    param (
        [Parameter(Mandatory=$true)]
        [string]$SkuPartNumber,
        
        [Parameter(Mandatory=$true)]
        $LicenseDetails
    )
    
    # Check cache first
    if ($script:LicenseMappingCache.ContainsKey($SkuPartNumber)) {
        return $script:LicenseMappingCache[$SkuPartNumber]
    }
    
    # If not in cache, look up the name
    $licenseName = switch ($SkuPartNumber) {
        # Microsoft 365 and Office 365 Plans
        "O365_BUSINESS_ESSENTIALS" { "Microsoft 365 Business Basic" }
        "O365_BUSINESS_PREMIUM" { "Microsoft 365 Business Standard" }
        "O365_BUSINESS" { "Microsoft 365 Apps for Business" }
        "O365_ENTERPRISE" { "Microsoft 365 Apps for Enterprise" }
        "ENTERPRISEPACK" { "Office 365 E3" }
        "ENTERPRISEPREMIUM" { "Office 365 E5" }
        "STANDARDPACK" { "Office 365 E2" }
        "DESKLESSPACK" { "Office 365 F1" }
        "M365_F1" { "Microsoft 365 F3" }
        "M365_E3" { "Microsoft 365 E3" }
        "M365_E5" { "Microsoft 365 E5" }
        "M365_BUSINESS_PREMIUM" { "Microsoft 365 Business Premium" }
        "M365_BUSINESS_STANDARD" { "Microsoft 365 Business Standard" }
        "M365_BUSINESS_BASIC" { "Microsoft 365 Business Basic" }
        "M365_ENTERPRISE" { "Microsoft 365 Enterprise" }
        "M365_F3" { "Microsoft 365 F3" }
        "M365_F5" { "Microsoft 365 F5" }
        "M365_G3" { "Microsoft 365 G3" }
        "M365_G5" { "Microsoft 365 G5" }
        "M365_E1" { "Microsoft 365 E1" }
        "M365_E2" { "Microsoft 365 E2" }
        "M365_E4" { "Microsoft 365 E4" }
        "O365_E1" { "Office 365 E1" }
        "O365_E2" { "Office 365 E2" }
        "O365_E4" { "Office 365 E4" }
        "O365_F3" { "Office 365 F3" }
        "O365_G3" { "Office 365 G3" }
        "O365_G5" { "Office 365 G5" }
        
        # Exchange Online Plans
        "EXCHANGESTANDARD" { "Exchange Online (Plan 1)" }
        "EXCHANGEENTERPRISE" { "Exchange Online (Plan 2)" }
        "EXCHANGEARCHIVE_ADDON" { "Exchange Online Archiving" }
        "EXCHANGEONLINE" { "Exchange Online" }
        "EXCHANGE_S_ESSENTIALS" { "Exchange Online Essentials" }
        "EXCHANGE_S_DESKLESS" { "Exchange Online Kiosk" }
        "EXCHANGE_S_STANDARD" { "Exchange Online (Plan 1)" }
        "EXCHANGE_S_ENTERPRISE" { "Exchange Online (Plan 2)" }
        
        # SharePoint Online Plans
        "SHAREPOINTSTANDARD" { "SharePoint Online (Plan 1)" }
        "SHAREPOINTENTERPRISE" { "SharePoint Online (Plan 2)" }
        "SHAREPOINTWAC" { "Office Online" }
        "SHAREPOINT_S_ESSENTIALS" { "SharePoint Online Essentials" }
        "SHAREPOINT_S_DESKLESS" { "SharePoint Online Kiosk" }
        "SHAREPOINT_S_STANDARD" { "SharePoint Online (Plan 1)" }
        "SHAREPOINT_S_ENTERPRISE" { "SharePoint Online (Plan 2)" }
        
        # Teams Plans
        "TEAMS_EXPLORATORY" { "Microsoft Teams Exploratory" }
        "TEAMS1" { "Microsoft Teams (Free)" }
        "TEAMS_COMMERCIAL_TRIAL" { "Microsoft Teams Trial" }
        "MCOMEETADV" { "Microsoft Teams Audio Conferencing" }
        "MCOPSTN2" { "Microsoft 365 Business Voice" }
        "MCOPSTN1" { "Microsoft 365 Domestic Calling Plan" }
        "MCOPSTN5" { "Microsoft 365 International Calling Plan" }
        "TEAMS_EXPLORATORY" { "Microsoft Teams Exploratory" }
        "TEAMS_FREE" { "Microsoft Teams (Free)" }
        "TEAMS_ESSENTIALS" { "Microsoft Teams Essentials" }
        "TEAMS_STANDARD" { "Microsoft Teams Standard" }
        "TEAMS_ENTERPRISE" { "Microsoft Teams Enterprise" }
        
        # Security and Compliance
        "EMS" { "Enterprise Mobility + Security E3" }
        "EMSPREMIUM" { "Enterprise Mobility + Security E5" }
        "IDENTITY_THREAT_PROTECTION" { "Microsoft Defender for Office 365 (Plan 1)" }
        "IDENTITY_THREAT_PROTECTION_FOR_EMS_E5" { "Microsoft Defender for Office 365 (Plan 2)" }
        "ATP_ENTERPRISE" { "Microsoft Defender for Office 365 (Plan 1)" }
        "DEFENDER_ENDPOINT_P1" { "Microsoft Defender for Endpoint P1" }
        "DEFENDER_ENDPOINT_P2" { "Microsoft Defender for Endpoint P2" }
        "DEFENDER_OFFICE_365_P1" { "Microsoft Defender for Office 365 (Plan 1)" }
        "DEFENDER_OFFICE_365_P2" { "Microsoft Defender for Office 365 (Plan 2)" }
        "DEFENDER_IDENTITY" { "Microsoft Defender for Identity" }
        "DEFENDER_CLOUD_APPS" { "Microsoft Defender for Cloud Apps" }
        "DEFENDER_VULNERABILITY_MANAGEMENT" { "Microsoft Defender Vulnerability Management" }
        
        # Power Platform
        "POWER_BI_STANDARD" { "Power BI (Free)" }
        "POWER_BI_PRO" { "Power BI Pro" }
        "POWER_BI_PREMIUM" { "Power BI Premium" }
        "POWER_BI_PREMIUM_PER_USER" { "Power BI Premium Per User" }
        "FLOW_FREE" { "Power Automate Free" }
        "POWERAPPS_VIRAL" { "Power Apps Free" }
        "POWERAPPS_PER_USER" { "Power Apps per user plan" }
        "POWERAUTOMATE_ATTENDED_RPA" { "Power Automate per user plan with attended RPA" }
        "POWERAUTOMATE_UNATTENDED_RPA" { "Power Automate per user plan with unattended RPA" }
        "POWER_VIRTUAL_AGENTS" { "Power Virtual Agents" }
        
        # Windows 365
        "WINDOWS_365_BUSINESS" { "Windows 365 Business" }
        "WINDOWS_365_ENTERPRISE" { "Windows 365 Enterprise" }
        "WINDOWS_365_F1" { "Windows 365 F1" }
        "WINDOWS_365_F3" { "Windows 365 F3" }
        "WINDOWS_365_E3" { "Windows 365 E3" }
        "WINDOWS_365_E5" { "Windows 365 E5" }
        
        # New mapping
        "SPB" { "Microsoft 365 Business Premium" }
        
        # Default case - return the original SKU if no mapping exists
        default { 
            # Try to get the service plan name from the license details
            $servicePlan = $license.ServicePlans | Where-Object { $_.ServicePlanId -eq $license.SkuId } | Select-Object -First 1
            if ($servicePlan) {
                $servicePlan.ServicePlanName
            } else {
                $license.SkuPartNumber
            }
        }
    }
    
    # Cache the result
    $script:LicenseMappingCache[$SkuPartNumber] = $licenseName
    return $licenseName
}

function Export-UserLicenses {
    param (
        [Parameter(Mandatory=$true)]
        [string]$UserPrincipalName,
        
        [Parameter(Mandatory=$false)]
        [string]$ExportPath = (Join-Path $PSScriptRoot "exports"),
        
        [Parameter(Mandatory=$false)]
        [string]$ExistingCsvPath = $null
    )
    
    try {
        Write-Host "Exporting licenses for $UserPrincipalName..." -ForegroundColor Cyan
        
        # Ensure the export directory exists
        if (-not (Test-Path $ExportPath)) {
            New-Item -ItemType Directory -Path $ExportPath | Out-Null
        }
        
        $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
        
        # If an existing CSV is provided, we'll append to it, otherwise create a new one
        $exportFile = if ($ExistingCsvPath -and (Test-Path $ExistingCsvPath)) {
            $ExistingCsvPath
        } else {
            $upnForFilename = $UserPrincipalName
            Join-Path $ExportPath "$upnForFilename-Licenses-$timestamp.csv"
        }
        
        # Get user details
        $user = Get-MgUser -UserId $UserPrincipalName -ErrorAction Stop
        $userId = $user.Id
        
        Write-Host "Found user with ID: $userId" -ForegroundColor Cyan
        
        # Get licenses
        Write-Host "Getting assigned licenses..." -ForegroundColor Cyan
        $licenses = @()
        try {
            $licenseDetails = Get-MgUserLicenseDetail -UserId $userId -ErrorAction Stop
            foreach ($license in $licenseDetails) {
                $licenseName = Get-LicenseFriendlyName -SkuPartNumber $license.SkuPartNumber -LicenseDetails $license
                
                $licenses += [PSCustomObject]@{
                    "License Name" = $licenseName
                    "SKU Part Number" = $license.SkuPartNumber
                }
            }
        }
        catch {
            Write-Host "Warning: Could not retrieve licenses: $_" -ForegroundColor Yellow
        }
        
        # Export licenses to CSV
        if ($licenses.Count -gt 0) {
            if (Test-Path $exportFile) {
                Add-Content -Path $exportFile -Value "`n`n# User Licenses`n"
                $licenses | ConvertTo-Csv -NoTypeInformation | Select-Object -Skip 1 | Add-Content -Path $exportFile
                Write-Host "✓ Successfully appended $($licenses.Count) licenses to: $exportFile" -ForegroundColor Green
            } else {
                $licenses | Export-Csv -Path $exportFile -NoTypeInformation
                Write-Host "✓ Successfully exported $($licenses.Count) licenses to: $exportFile" -ForegroundColor Green
            }
        }
        else {
            Write-Host "No licenses found for $UserPrincipalName" -ForegroundColor Yellow
            if (-not (Test-Path $exportFile)) {
                [PSCustomObject]@{
                    "License Name" = ""
                    "SKU Part Number" = ""
                } | Export-Csv -Path $exportFile -NoTypeInformation
            } else {
                Add-Content -Path $exportFile -Value "`n`n# User Licenses`n"
                Add-Content -Path $exportFile -Value """License Name"",""SKU Part Number""`n"","""
            }
        }
        
        return $exportFile
    }
    catch {
        Write-Host "✗ Failed to export licenses for $UserPrincipalName" -ForegroundColor Red
        Write-Host "Error: $_" -ForegroundColor Red
        return $null
    }
}
#endregion

#region CSV Processing Functions
function Read-OffboardingCSV {
    param (
        [Parameter(Mandatory=$false)]
        [string]$CSVPath = (Join-Path $PSScriptRoot "Term_User.csv")
    )
    
    try {
        # Check if file exists
        if (-not (Test-Path $CSVPath)) {
            Write-Host "CSV file not found: $CSVPath" -ForegroundColor Red
            return $null
        }
        
        # Import CSV
        $csvData = Import-Csv -Path $CSVPath
        
        # Validate CSV structure (needs required columns)
        $requiredColumns = @("Term_User_UPN", "Delegate1", "OOO")
        $hasRequiredColumns = $true
        
        foreach ($column in $requiredColumns) {
            if (-not $csvData[0].PSObject.Properties.Name.Contains($column)) {
                Write-Host "CSV is missing required column: $column" -ForegroundColor Red
                $hasRequiredColumns = $false
            }
        }
        
        if (-not $hasRequiredColumns) {
            Write-Host "Please ensure your CSV has the following columns:" -ForegroundColor Yellow
            Write-Host "Term_User_UPN, Delegate1, Delegate2, Delegate3, OOO" -ForegroundColor Yellow
            return $null
        }
        
        Write-Host "Successfully loaded offboarding data for $($csvData.Count) users" -ForegroundColor Green
        return $csvData
    }
    catch {
        Write-Host "Error reading CSV file: $_" -ForegroundColor Red
        return $null
    }
}

function Create-CSVTemplate {
    param (
        [Parameter(Mandatory=$false)]
        [string]$OutputPath = (Join-Path $PSScriptRoot "Term_User.csv")
    )
    
    try {
        $template = @(
            [PSCustomObject]@{
                Term_User_UPN = "user@domain.com"
                Delegate1 = "delegate1@domain.com"
                Delegate2 = "delegate2@domain.com"
                Delegate3 = "delegate3@domain.com"
                OOO = "I am no longer with the company. Please contact [Name] at [Email] for assistance."
            }
        )
        
        $template | Export-Csv -Path $OutputPath -NoTypeInformation
        Write-Host "CSV template created at: $OutputPath" -ForegroundColor Green
        return $OutputPath
    }
    catch {
        Write-Host "Error creating CSV template: $_" -ForegroundColor Red
        return $null
    }
}
#endregion

#region Mailbox Management Functions
function Add-MailboxDelegates {
    param (
        [Parameter(Mandatory=$true)]
        [string]$SharedMailboxUPN,
        
        [Parameter(Mandatory=$false)]
        [string]$Delegate1,
        
        [Parameter(Mandatory=$false)]
        [string]$Delegate2,
        
        [Parameter(Mandatory=$false)]
        [string]$Delegate3
    )
    
    $successCount = 0
    $delegates = @($Delegate1, $Delegate2, $Delegate3) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    
    if ($delegates.Count -eq 0) {
        Write-Host "No delegates specified for $SharedMailboxUPN" -ForegroundColor Yellow
        return $false
    }
    
    try {
        Write-Host "Adding delegates to mailbox $SharedMailboxUPN..." -ForegroundColor Cyan
        
        # Check if we're connected to Exchange Online
        try {
            $null = Get-Mailbox -ResultSize 1 -ErrorAction Stop
        }
        catch {
            Write-Host "Not connected to Exchange Online. Attempting to connect..." -ForegroundColor Yellow
            if (!(Connect-ToExchangeOnline)) {
                throw "Unable to connect to Exchange Online. Cannot add delegates."
            }
        }
        
        # Add full access permissions for each delegate
        foreach ($delegate in $delegates) {
            try {
                Write-Host "Adding full access for $delegate..." -ForegroundColor Cyan
                Add-MailboxPermission -Identity $SharedMailboxUPN -User $delegate -AccessRights FullAccess -InheritanceType All -AutoMapping $true -ErrorAction Stop
                
                # Removed SendAs permission granting as requested
                
                Write-Host "✓ Successfully added $delegate as a delegate with Full Access permissions" -ForegroundColor Green
                $successCount++
            }
            catch {
                Write-Host "✗ Failed to add delegate $delegate to $SharedMailboxUPN" -ForegroundColor Red
                Write-Host "Error: $_" -ForegroundColor Red
            }
        }
        
        return ($successCount -gt 0)
    }
    catch {
        Write-Host "✗ Failed to add delegates to $SharedMailboxUPN" -ForegroundColor Red
        Write-Host "Error: $_" -ForegroundColor Red
        return $false
    }
}

function Set-OutOfOfficeMessage {
    param (
        [Parameter(Mandatory=$true)]
        [string]$UserPrincipalName,
        
        [Parameter(Mandatory=$true)]
        [string]$Message
    )
    
    try {
        Write-Host "Setting out of office message for $UserPrincipalName..." -ForegroundColor Cyan
        
        # Set automatic replies (out of office)
        Set-MailboxAutoReplyConfiguration -Identity $UserPrincipalName -AutoReplyState Enabled -InternalMessage $Message -ExternalMessage $Message -ErrorAction Stop
        
        Write-Host "✓ Successfully set out of office message for $UserPrincipalName" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "✗ Failed to set out of office message for $UserPrincipalName" -ForegroundColor Red
        Write-Host "Error: $_" -ForegroundColor Red
        return $false
    }
}
#endregion

#region Main Function
function Start-UserTermination {
    param (
        [Parameter(Mandatory=$true)]
        [string]$UserPrincipalName,
        
        [Parameter(Mandatory=$false)]
        [string]$Delegate1,
        
        [Parameter(Mandatory=$false)]
        [string]$Delegate2,
        
        [Parameter(Mandatory=$false)]
        [string]$Delegate3,
        
        [Parameter(Mandatory=$false)]
        [string]$OutOfOfficeMessage,
        
        [Parameter(Mandatory=$false)]
        [switch]$SkipConvertToShared,
        
        [Parameter(Mandatory=$false)]
        [switch]$SkipExport
    )
    
    # Step 1: Check if user exists
    Write-Host "`n==== Starting termination process for $UserPrincipalName ====`n" -ForegroundColor Cyan
    
    try {
        Write-Host "Checking if user $UserPrincipalName exists..." -ForegroundColor Cyan
        $user = Get-MgUser -UserId $UserPrincipalName -ErrorAction Stop
        Write-Host "✓ User found: $($user.DisplayName)" -ForegroundColor Green
        
        # Step A: Export user information if not skipped
        if (-not $SkipExport) {
            # Export group memberships
            $exportFile = Export-UserMemberships -UserPrincipalName $UserPrincipalName
            
            # Export license information
            if ($exportFile) {
                Export-UserLicenses -UserPrincipalName $UserPrincipalName -ExistingCsvPath $exportFile
            }
        }
        
        # Step 3: Perform termination actions
        $results = @{}
        
        # Block sign-in
        $results.SignInBlocked = Block-UserSignIn -UserPrincipalName $UserPrincipalName
        
        # Revoke active sessions
        $results.SessionsRevoked = Revoke-UserSessions -UserPrincipalName $UserPrincipalName
        
        # Rename display name
        $results.DisplayNameChanged = Rename-UserDisplayName -UserPrincipalName $UserPrincipalName
        
        # Exchange Online operations
        # Hide from GAL
        $results.HiddenFromGAL = Hide-UserFromGAL -UserPrincipalName $UserPrincipalName
        
        # Set out of office message if provided and not empty
        if (-not [string]::IsNullOrWhiteSpace($OutOfOfficeMessage)) {
            $results.OutOfOfficeSet = Set-OutOfOfficeMessage -UserPrincipalName $UserPrincipalName -Message $OutOfOfficeMessage
        }
        else {
            $results.OutOfOfficeSet = $null
            Write-Host "Out of Office message not set (blank in CSV)" -ForegroundColor Yellow
        }
        
        # Convert to shared mailbox if not skipped
        if (-not $SkipConvertToShared) {
            $results.ConvertedToShared = Convert-ToSharedMailbox -UserPrincipalName $UserPrincipalName
            
            # Add delegates if mailbox is converted to shared and delegates are provided
            if ($results.ConvertedToShared -and 
                (-not [string]::IsNullOrWhiteSpace($Delegate1) -or 
                 -not [string]::IsNullOrWhiteSpace($Delegate2) -or 
                 -not [string]::IsNullOrWhiteSpace($Delegate3))) {
                $results.DelegatesAdded = Add-MailboxDelegates -SharedMailboxUPN $UserPrincipalName -Delegate1 $Delegate1 -Delegate2 $Delegate2 -Delegate3 $Delegate3
            }
        }
        else {
            $results.ConvertedToShared = $false
            $results.DelegatesAdded = $false
        }
        
        # Step 4: Summarize results
        Write-Host "`n==== Termination Results for $UserPrincipalName ====`n" -ForegroundColor Cyan
        
        Write-Host "Sign-In Blocked: $($results.SignInBlocked ? '✓' : '✗')" -ForegroundColor ($results.SignInBlocked ? 'Green' : 'Red')
        Write-Host "Sessions Revoked: $($results.SessionsRevoked ? '✓' : '✗')" -ForegroundColor ($results.SessionsRevoked ? 'Green' : 'Red')
        Write-Host "Display Name Changed: $($results.DisplayNameChanged ? '✓' : '✗')" -ForegroundColor ($results.DisplayNameChanged ? 'Green' : 'Red')
        Write-Host "Hidden from GAL: $($results.HiddenFromGAL ? '✓' : '✗')" -ForegroundColor ($results.HiddenFromGAL ? 'Green' : 'Red')
        
        if (-not [string]::IsNullOrWhiteSpace($OutOfOfficeMessage)) {
            Write-Host "Out of Office Set: $($results.OutOfOfficeSet ? '✓' : '✗')" -ForegroundColor ($results.OutOfOfficeSet ? 'Green' : 'Red')
        }
        
        if (-not $SkipConvertToShared) {
            Write-Host "Converted to Shared: $($results.ConvertedToShared ? '✓' : '✗')" -ForegroundColor ($results.ConvertedToShared ? 'Green' : 'Red')
            
            if ($results.ConvertedToShared -and 
                (-not [string]::IsNullOrWhiteSpace($Delegate1) -or 
                 -not [string]::IsNullOrWhiteSpace($Delegate2) -or 
                 -not [string]::IsNullOrWhiteSpace($Delegate3))) {
                Write-Host "Delegates Added: $($results.DelegatesAdded ? '✓' : '✗')" -ForegroundColor ($results.DelegatesAdded ? 'Green' : 'Red')
            }
        }
        
        if (-not $SkipExport) {
            Write-Host "User Data Exported: $($exportFile ? '✓' : '✗')" -ForegroundColor ($exportFile ? 'Green' : 'Red')
            if ($exportFile) {
                Write-Host "Export file: $exportFile" -ForegroundColor Cyan
            }
        }
        
        Write-Host "`nTermination process completed for $UserPrincipalName." -ForegroundColor Cyan
        
        $finalUser = Get-MgUser -UserId $UserPrincipalName -ErrorAction Stop
        $results.SignInBlocked = ($finalUser.AccountEnabled -eq $false -or $finalUser.AccountEnabled -eq "False")
        
        return $true
    }
    catch {
        Write-Host "✗ User not found or cannot be accessed: $UserPrincipalName" -ForegroundColor Red
        Write-Host "Error: $_" -ForegroundColor Red
        return $false
    }
}

function Main {
    # Test mode for sign-in blocking
    if ($args -contains "-TestBlock") {
        Write-Host "Running in test mode for sign-in blocking..." -ForegroundColor Cyan
        
        # Establish connections
        Write-Host "Establishing required connections..." -ForegroundColor Cyan
        
        # Connect to Microsoft Graph
        $connectedToGraph = Connect-ToGraph
        if (!$connectedToGraph) {
            Write-Host "✗ Failed to connect to Microsoft Graph. Cannot proceed with test." -ForegroundColor Red
            return
        }
        
        # Get user to test
        $testUser = Read-Host "Enter the UPN of the user to test blocking (e.g., user@domain.com)"
        
        Write-Host "`nTesting sign-in blocking for $testUser..." -ForegroundColor Cyan
        
        # Test the blocking
        $result = Block-UserSignIn -UserPrincipalName $testUser
        
        if ($result) {
            Write-Host "`n✓ Test completed successfully. User sign-in should be blocked." -ForegroundColor Green
        } else {
            Write-Host "`n✗ Test failed. User sign-in may not be blocked." -ForegroundColor Red
        }
        
        return
    }

    # Main script execution - processing users from CSV
    $csvPath = Join-Path $PSScriptRoot "Term_User.csv"

    # Establish connections at the start
    Write-Host "Establishing required connections..." -ForegroundColor Cyan

    # Connect to Microsoft Graph
    $connectedToGraph = Connect-ToGraph
    if (!$connectedToGraph) {
        Write-Host "✗ Failed to connect to Microsoft Graph. Cannot proceed with termination." -ForegroundColor Red
        return
    }

    # Connect to Exchange Online
    $connectedToExchange = Connect-ToExchangeOnline
    if (!$connectedToExchange) {
        Write-Host "✗ Failed to connect to Exchange Online. Cannot proceed with termination." -ForegroundColor Red
        return
    }

    Write-Host "✓ All required connections established successfully" -ForegroundColor Green
        
        # Check if a template CSV should be created
    if ($args -contains "-CreateTemplate") {
            Write-Host "Creating CSV template file..." -ForegroundColor Cyan
            $templatePath = Create-CSVTemplate
            
            if ($templatePath) {
                Write-Host "`nTemplate created. Please edit this file and run the script again." -ForegroundColor Green
                
                # Try to open the template file
                try {
                    Invoke-Item $templatePath
                }
                catch {
                    Write-Host "Could not automatically open the CSV file. Please open it manually at: $templatePath" -ForegroundColor Yellow
                }
            }
        return
        }
        
        # Read the CSV data
        Write-Host "Reading user data from Term_User.csv..." -ForegroundColor Cyan
        $offboardingData = Read-OffboardingCSV -CSVPath $csvPath
        
        if (-not $offboardingData) {
            # If the CSV doesn't exist or has invalid format, offer to create a template
            Write-Host "`nNo valid CSV data found. Would you like to create a template CSV file? (Y/N)" -ForegroundColor Yellow
            $createTemplate = Read-Host
            
            if ($createTemplate -eq "Y" -or $createTemplate -eq "y") {
                $templatePath = Create-CSVTemplate
                
                if ($templatePath) {
                    Write-Host "`nTemplate created. Please edit this file and run the script again." -ForegroundColor Green
                    
                    # Try to open the template file
                    try {
                        Invoke-Item $templatePath
                    }
                    catch {
                        Write-Host "Could not automatically open the CSV file. Please open it manually at: $templatePath" -ForegroundColor Yellow
                    }
                }
            }
        return
        }
        
        # Process each user in the CSV
        $totalUsers = $offboardingData.Count
        $currentUser = 0
        
        foreach ($userData in $offboardingData) {
            $currentUser++
            Write-Host "`n[$currentUser/$totalUsers] Processing user: $($userData.Term_User_UPN)" -ForegroundColor Cyan
            
            # Start the termination process for this user
            Start-UserTermination `
                -UserPrincipalName $userData.Term_User_UPN `
                -Delegate1 $userData.Delegate1 `
                -Delegate2 $userData.Delegate2 `
                -Delegate3 $userData.Delegate3 `
                -OutOfOfficeMessage $userData.OOO
        }
        
        Write-Host "`nAll users processed. Termination script completed." -ForegroundColor Green
}


# --- Script Entry Point ---
# Temporarily suppress warnings to hide benign module loading messages
$OldWarningPreference = $WarningPreference
$WarningPreference = 'SilentlyContinue'

try {
    # Run the main script logic
    Main
}
finally {
    # This block always runs, ensuring disconnection and restoration of settings
    Write-Host "`nDisconnecting from services and restoring settings..." -ForegroundColor Yellow
    Disconnect-FromGraph
    Disconnect-FromExchangeOnline 
    $WarningPreference = $OldWarningPreference
} 