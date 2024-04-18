# Export a W.Sever IIS (FTP) Sites' Virtual Directories and their Physical Paths to CSV
# Ensure the module WebAdministration installed and is imported
#--------------------------------------------------------------------------------------------------------------

# Variables
$csvpath = "C:\users\$env:username\desktop\IISoutput.csv"

# Get all IIS sites
$Websites = Get-ChildItem IIS:\\Sites

# Loop through each site
foreach ($Site in $Websites) {
    $webapps = Get-WebApplication -Site $Site.name
    $VDir = Get-WebVirtualDirectory -Site $Site.name

    foreach ($webvdirectory in $VDir) {
        $vDirName = $webvdirectory.path
        $physicalPath = $webvdirectory.PhysicalPath

        # Create a custom object with relevant information
        $vDirInfo = [PSCustomObject]@{
            Site_Name = $Site.name
            Virtual_Directory_Name = $vDirName
            Physical_Path = $physicalPath
        }

        # Export the object to a CSV file
        $vDirInfo | Export-Csv -Path "$csvpath" -Append -NoTypeInformation
    }
}