# Function to deploy extension for a specific browser
function Deploy-Extension {
    param (
        [string]$BrowserName,
        [string]$RegistryPath,
        [string]$ExtensionID
    )
 
    # Create the registry key if it doesn't exist
    if (-not (Test-Path $RegistryPath)) {
        New-Item -Path $RegistryPath -Force | Out-Null
    }
 
    # Set the extension ID in the registry
    New-ItemProperty -Path $RegistryPath -Name "1" -Value $ExtensionID -PropertyType String -Force | Out-Null
 
    Write-Host "uBlock Origin deployed for $BrowserName"
}
 
# uBlock Origin Extension ID
$uBlockOriginID_Chrome = "cjpalhdlnbpafiamejdnhcphjbkeiagm"
$uBlockOriginID_Edge = "odfafepnkmbhccpbejgmiehpchacaeak"
 
# Deploy for Chrome
$ChromeRegistryPath = "HKLM:\SOFTWARE\Policies\Google\Chrome\ExtensionInstallForcelist"
Deploy-Extension -BrowserName "Chrome" -RegistryPath $ChromeRegistryPath -ExtensionID $uBlockOriginID_Chrome
 
# Deploy for Edge
$EdgeRegistryPath = "HKLM:\SOFTWARE\Policies\Microsoft\Edge\ExtensionInstallForcelist"
Deploy-Extension -BrowserName "Edge" -RegistryPath $EdgeRegistryPath -ExtensionID $uBlockOriginID_Edge
 
Write-Host "Deployment complete. Please restart browsers for changes to take effect."
