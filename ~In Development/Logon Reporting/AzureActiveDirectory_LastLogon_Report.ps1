# If you cant connect to AzureADPreview\Connect-AzureAD - uncomment the "Install-Module AzureADPreview –AllowClobber" command and try again
#Install-Module AzureADPreview –AllowClobber

#Connect/Signin to AzureAD
AzureADPreview\Connect-AzureAD

# Fetches the last month's Azure Active Directory sign-in data
# Azure will only let us pull logs for the last 30 days. You can adjust the "AddDays(-30) below to anywhere between 01-30.

# Start measuring time
$StartTime = Get-Date

CLS; $StartDate = (Get-Date).AddDays(-30); $StartDate = Get-Date($StartDate) -format yyyy-MM-dd
Write-Host "Fetching data from Azure Active Directory..."
$Records = Get-AzureADAuditSignInLogs -Filter "createdDateTime gt $StartDate" -all:$True
$Report = [System.Collections.Generic.List[Object]]::new()
ForEach ($Rec in $Records) {
    Switch ($Rec.Status.ErrorCode) {
      "0" {$Status = "Success"}
      default {$Status = $Rec.Status.FailureReason}
    }
    $ReportLine = [PSCustomObject] @{
           TimeStamp   = Get-Date($Rec.CreatedDateTime) -format g
           User        = $Rec.UserPrincipalName
           Name        = $Rec.UserDisplayName
           IPAddress   = $Rec.IpAddress
           ClientApp   = $Rec.ClientAppUsed
           Device      = $Rec.DeviceDetail.OperatingSystem
           Location    = $Rec.Location.City + ", " + $Rec.Location.State + ", " + $Rec.Location.CountryOrRegion
           Appname     = $Rec.AppDisplayName
           Resource    = $Rec.ResourceDisplayName
           Status      = $Status
           Correlation = $Rec.CorrelationId
           Interactive = $Rec.IsInteractive }
      $Report.Add($ReportLine)
}

Write-Host $Report.Count "sign-in audit records processed."

# Export the data to a CSV file
$logdate = get-date -f MM-dd-yyyy
$Report | Export-Csv -Path "c:\users\$env:username\desktop\AzureActiveDirectory_LastLogon_Report-$logdate.csv" -NoTypeInformation

# Stop measuring time
$EndTime = Get-Date

# Calculate and display the script runtime
$ElapsedTime = $EndTime - $StartTime
Write-Host "The script ran for $($ElapsedTime.TotalMinutes) minutes."
