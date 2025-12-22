# Variables
$DhcpServer = "" # Replace with the name or IP address of your DHCP server
$inputMAC = Read-Host "Enter the MAC Address (without - or : )"

# Transform '$inputMAC' to add ' - ' evert 2 characters before searching
$macaddress = $inputMAC -replace '(.{2})', '$1-'
$macaddress = $macaddress.TrimEnd('-')
#----------------------------------------------------------------------------------------
# Script Action: Lookup DHCP Lease/Reservation by MAC Address
Get-DhcpServerv4Scope -ComputerName $dhcpserver | Get-DhcpServerv4Lease -ComputerName $dhcpserver | Where-Object{$_.clientid -eq $macaddress}