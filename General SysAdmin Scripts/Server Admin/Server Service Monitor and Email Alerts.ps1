<#
1: Install the CredentialManager Module so you don't store passwords in your script
Install-Module -Name CredentialManager -Scope AllUsers

2: Create the Stored Credential with a Target Name and input the username and password
# Use Set-StoredCredential to save the credentials
New-StoredCredential -Target "YourTargetName" -UserName "email@contoso.com" -Password "YourPasswordHere"

3: Use the credential in your script
# Retrieve the stored credential to use in a Script
$cred = Get-StoredCredential -Target "YourTargetName"
--------------------------------------------------------------------------------------------------------------------------------------------------------
(4:) If you need to review the password, enter the following command to display what you have the password configured to

# Convert the SecureString password to plain text to display/review
$passwordPlainText = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($cred.Password))

# Display the password
$passwordPlainText


• This is an example script that checks if IISADMIN and the SMTP Relay Service is running. If it is off, it will try and turn it on and send an email
 and report that the service was off, but now is on. If it faied to turn it on, it will send a different email.

• You must make sure the account you're using is authorized to send emails with SMTP.

• After you have the script - use Task Scheduler to run this script at your preferred interval
• You should run the script as the SYSTEM user so use PSEXEC to install and configure the Credentail Manager Module in Powershell so it stores the creds correctly and uses them in the Task Scheduler
#>

# Define the services to monitor
$services = @("IISADMIN", "SMTPSVC") # add/remove whatever else you want

# Email settings
$emailSmtpServer = "" # example: "smtp.office365.com"
$emailSmtpPort = "" # example: "587"
$emailFrom = "" # example: "smtpalert@contoso.com"
$emailTo = "" # example: "sysadmins@contoso.com"
$emailCredentialUser = "" # account used to auth with the SMTP relay -- example: "smtprelayaccount@contoso.com"
$credential = Get-StoredCredential -Target "PS-SMTP-Credentials" # or whatever name you called the stored creds

#--------------------------------------------------------------------------------------------------------------------------------------------------------

$restartLog = @()
$debugLog = @()

foreach ($service in $services) {
    $serviceStatus = Get-Service -Name $service
    $debugLog += "Service $service status: $($serviceStatus.Status)"

    if ($serviceStatus.Status -ne "Running") {
        # Try restarting the service
        try {
            Restart-Service -Name $service -ErrorAction Stop

            # Check if the service is running after restart
            $serviceStatusAfterRestart = Get-Service -Name $service

            if ($serviceStatusAfterRestart.Status -eq "Running") {
                # Add log indicating restart was successful
                $restartLog += "The $service service was restarted and is now running again."
            } else {
                # Add log indicating service could not be started after a restart attempt
                $restartLog += "Attempted to restart the $service service, but it is still not running."
            }
        } catch {
            # Service restart failed, add to the log
            $restartLog += "An error occurred while trying to restart the service $service. Error Details: " + $_.Exception.Message
        }
    }
}

# Combine debug and restart logs
$combinedLog = $debugLog + $restartLog

# Send a consolidated email if there's any log entry
if ($restartLog.Count -gt 0) {
    $body = $combinedLog -join "`n`n"
    Send-MailMessage -SmtpServer $emailSmtpServer -Port $emailSmtpPort -UseSsl -From $emailFrom -To $emailTo -Credential $credential -Subject "SERVER01: Service Restart Report" -Body $body
}