# Define the registry paths to search for installed applications
$uninstallKeys = @(
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*",
    "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*"
)

# Iterate through each registry path
foreach ($key in $uninstallKeys) {
    # Get the properties of each subkey
    Get-ItemProperty $key | Where-Object {
        $_.DisplayName -like "*Adobe Acrobat*"
    } | ForEach-Object {
        $displayName = $_.DisplayName
        $uninstallString = $_.UninstallString

        # Display the application name and uninstall string
        Write-Host "Found Adobe product: $displayName"
        Write-Host "Uninstall String: $uninstallString"

        # Modify the uninstall string to use /X instead of /I
        if ($uninstallString -match "MsiExec.exe /I") {
            $uninstallString = $uninstallString -replace "/I", "/X"
        }

        # Execute the uninstall command
        if ($uninstallString) {
            Write-Host "Uninstalling $displayName..."
            Start-Process -FilePath "cmd.exe" -ArgumentList "/c $uninstallString /qn /norestart" -Wait
            Write-Host "$displayName has been uninstalled."
        } else {
            Write-Host "No uninstall string found for $displayName."
        }
        Write-Host "`n"
    }
}