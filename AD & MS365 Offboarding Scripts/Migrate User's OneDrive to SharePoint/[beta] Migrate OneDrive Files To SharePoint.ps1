# Written by dangitbobby10
# This is still a prototype even though it works. I wanted to dress it up some more but also wanted to publish this as I have friends with a need for it.
# I know AZCopy is a thing - I'll work on a version that uses AZCopy at a later time.
#---------------------------------------------------------------------------------------------------------------------------------------------------
#---------------------------------------------------------------------------------------------------------------------------------------------------
#   ■■■■■■■■■■■■■■■■■■■■■■■■■■
#   █ Key Defined Variables: █
#   ■■■■■■■■■■■■■■■■■■■■■■■■■■
# Authenticate with Certificate Variables
    $clientid = ''
    $cert_pfx_path = "C:\some path to your pfx cert\pnponlineautomatedscripts.pfx"
    $cert_pfx_pw = (ConvertTo-SecureString -AsPlainText 'pass123word' -Force) # replace 'pass123word' if your pfx password
    $tenant_domain = "contoso.com"

# SharePoint Site and Document Library Details
    $sharePointSiteUrl = "https://contoso.sharepoint.com/sites/SITECOLLECTION/"       # please ensure you have "/" at the end of the URL
    $libraryPath = "Shared Documents"
    $TEMP_folderpath = "General/WHEREVER/THE/SECURE/FOLDER/IS/LOCATED"        # DO NOT leave a "/" at the end of the path. This specifies the parent directory in the site collection where the offboarded users documents reside.
#---------------------------------------------------------------------------------------------------------------------------------------------------
#---------------------------------------------------------------------------------------------------------------------------------------------------
# Additional Variables
    $trimmed_domain = $tenant_domain.Split('.')[0]
    $basesharepointurl = "$trimmed_domain.sharepoint.com"
    $base_admin_sharepoint_url = "$trimmed_domain-admin.sharepoint.com"
#---------------------------------------------------------------------------------------------------------------------------------------------------
# Check if running as Administrator
    if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
        Write-Warning "You need to run this script as an Administrator. Restarting with elevation..."
        Start-Process pwsh -Verb RunAs -ArgumentList ('-NoProfile -ExecutionPolicy Bypass -File "{0}"' -f ($myinvocation.MyCommand.Definition))
        exit
    }

# Check if the PnP.PowerShell module is installed
    if (!(Get-Module -ListAvailable -Name PnP.PowerShell)) {
        Install-Module -Name PnP.PowerShell -Scope CurrentUser -Force
        Import-Module -Name PnP.PowerShell
    } else {
        Write-Host "PnP.PowerShell is already installed."
    }
#---------------------------------------------------------------------------------------------------------------------------------------------------
# Prompts for the offboarded user's email
    Connect-PnPOnline -Url $base_admin_sharepoint_url -ClientId $clientID -Tenant $tenant_domain -CertificatePath $cert_pfx_path -CertificatePassword $cert_pfx_pw
    function PromptForUserEmail {
        $userExists = $false
        $userProfileProperties = $null
        $promptfor_userEmail = $null  # Declare the variable outside the try-catch to widen its scope
        while (-not $userExists) {
            $promptfor_userEmail = Read-Host -Prompt 'Enter Email Address of user whose OneDrive data needs to be copied/moved to SharePoint'
            try {
                $userProfileProperties = Get-PnPUserProfileProperty -Account $promptfor_userEmail -ErrorAction Stop
                $userExists = $true
            } catch {
                Write-Host "Incorrect email inputted or user does not exist. Please try entering the email again." -ForegroundColor Red
            }
        }
        # Return both the user profile properties and the email address as a custom object
        return [PSCustomObject]@{
            Email = $promptfor_userEmail
            Properties = $userProfileProperties
        }
    }

# Use the function and store the returned values
    $result = PromptForUserEmail
    $userProfileProperties = $result.Properties
    $useroneDriveUrl = $userProfileProperties.PersonalUrl
    $sharepoint_useronedrive_root = $userProfileProperties.DisplayName -replace ' ', '_'
#---------------------------------------------------------------------------------------------------------------------------------------------------
# creates the folderpath that will be used during transfer
    $folderpath = "$TEMP_folderpath/$sharepoint_useronedrive_root"
#---------------------------------------------------------------------------------------------------------------------------------------------------
# Connect to the terminated user's OneDrive
    Connect-PnPOnline -Url $userOneDriveUrl -ClientId $clientID -Tenant $tenant_domain -CertificatePath $cert_pfx_path -CertificatePassword $cert_pfx_pw

# Enumerate the files in the user's OneDrive Documents library
    $oneDriveFiles = Get-PnPListItem -List "Documents" -Fields "FileLeafRef", "FileRef"

# Initialize an array to hold custom objects
    $fileInfoList = @()

    foreach ($file in $oneDriveFiles) {
        # Skip folders
        if ($file.FileSystemObjectType -eq "Folder") { continue }

        # Extract the relative path from FileRef and prepare the target SharePoint path
        $relativePath = $file["FileRef"].Substring($file["FileRef"].IndexOf('/Documents/') + 11)
        $targetFolderUrl = "$sharePointSiteUrl$libraryPath/$folderPath/$relativePath"
        
        # Remove file name to get folder path
        $targetFolderUrl = $targetFolderUrl -replace '/[^/]+$', ''

        # Create a custom object for each file
        $fileInfo = [PSCustomObject]@{
            SourceUrl      = $file["FileRef"]
            TargetFolderUrl = $targetFolderUrl
            FileName       = $file["FileLeafRef"]
            RelativePath   = $relativePath
        }

        # Add the custom object to the list
        $fileInfoList += $fileInfo
    }

# Now $fileInfoList contains all the information which can be used later
#---------------------------------------------------------------------------------------------------------------------------------------------------
# Reconnect to the SharePoint Site Collection
    Connect-PnPOnline -Url $sharePointSiteUrl -ClientId $clientID -Tenant $tenant_domain -CertificatePath $cert_pfx_path -CertificatePassword $cert_pfx_pw

    foreach ($info in $fileInfoList) {
        # Extract the SharePoint relative folder path from the TargetFolderUrl
            $relativeFolderPath = $info.TargetFolderUrl -replace [regex]::Escape($sharePointSiteUrl + $libraryPath + "/"), ''

        # Initialize the path to start from the library
            $currentPath = $libraryPath.TrimEnd('/')

        # Split the relative path to get individual folders
            $folders = $relativeFolderPath.Trim('/').Split('/')

        foreach ($folder in $folders) {
            # Update current path to include the next level of folder
                $currentPath = "$currentPath/$folder"
        
            # Convert SharePoint library path to server-relative URL for Get-PnPFolder
                $baseSharePointDomain = "https://$basesharepointurl"
                $serverRelativeUrl = $currentPath -replace [regex]::Escape($baseSharePointDomain), ''
        
            # Check if the folder exists
                try {
                    $folderExists = Get-PnPFolder -Url $serverRelativeUrl -ErrorAction Stop
                }
                catch {
                # Assuming the folder does not exist if an error is thrown
                    $folderExists = $null
                }
            
                if (-not $folderExists) {
                # Before the loop where you create folders
                    Write-Host "Verifying paths..."
                    Write-Host "Library Path: $libraryPath"
                    Write-Host "Folder Path: $folderPath"
        
                # Inside the loop, right before you attempt to create a folder
                    Write-Host "Attempting to create folder: $folder in $serverRelativeUrl"

                # Adjust the Add-PnPFolder command to ensure correct parent folder path
                    $parentFolderPath = Split-Path -Path $serverRelativeUrl -Parent

                # Ensure $parentFolderPath is a server-relative URL
                # Assuming $sharePointSiteUrl is something like 'https://contoso.sharepoint.com/sites/sitecollection/'
                # Trim the domain part from $serverRelativeUrl to make it server-relative
                    $siteRelativeUrl = $sharePointSiteUrl -replace [regex]::Escape($baseSharePointDomain), ''
                    $parentFolderPath = $parentFolderPath -replace [regex]::Escape($siteRelativeUrl), ''

                # Replace backslashes with forward slashes
                    $parentFolderPath = $parentFolderPath -replace '\\', '/'

                    Write-Host "Parent Folder for new folder: $parentFolderPath"

                # Assuming $parentFolderPath is the correct server-relative path to the parent folder
                    Add-PnPFolder -Name $folder -Folder $parentFolderPath -ErrorAction Stop
            }
        }
    }
    Write-Host "All required directories have been checked/created."
#---------------------------------------------------------------------------------------------------------------------------------------------------
# Connect back to user's OneDrive to begin data transfer
    Connect-PnPOnline -Url $userOneDriveUrl -ClientId $clientID -Tenant $tenant_domain -CertificatePath $cert_pfx_path -CertificatePassword $cert_pfx_pw

    foreach ($info in $fileInfoList) {
    # The source URL is the file's SharePoint URL
        $sourceUrl = $info.SourceUrl

    # The target URL is where the file needs to be copied to in SharePoint site
    # It's constructed from the SharePoint site URL, library path, and the file's relative path within the target folder structure
        $targetFileUrl = $info.TargetFolderUrl -replace [regex]::Escape($sharePointSiteUrl + $libraryPath + "/"), ''

    # Copy the file from OneDrive to the respective folder in SharePoint
    # Note: Copy-PnPFile operates from the context of the source, so ensure you're connected to OneDrive
        Write-Host "Copying file $($info.FileName) to $targetFileUrl"
        Copy-PnPFile -SourceUrl $sourceUrl -TargetUrl "$sharePointSiteUrl$libraryPath/$targetFileUrl" -OverwriteIfAlreadyExists -Force -ErrorAction Stop
    }

    Write-Host "File transfer complete."