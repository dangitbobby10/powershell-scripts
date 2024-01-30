Requirements:
1: Copy this script where AD services are installed (ideally a DC to whatever your primary AD server is)
2: Also ensure the 'LicenseFriendlyNamesScript' is also copied to the server as you'll need to define the path so it can be imported in.
3: Verify you can run an 'invoke-command' to the DC and AADSync server you define.
4: Confirm you have the following powershell modules installed:
	• MSOnline
	• ExchangeOnlineManagement
	• AzureAD
	
	(here's how to install them in case you dont have them:)
	Install-Module -Name MSOnline
	Install-Module -Name ExchangeOnlineManagement
	Install-Module -Name AzureAD



See line 36-57. Define the following variables:
• Define the domain controller to connect to (i advise targetting the DC that AADSync is targetting)
• Define the server with AADSync
• Define 'Disabled Users OU'
• Define 'LicenseFriendlyNamesScript' path for the MS365 Licenses. Reads as the actual license rather than the SKU.
• Define the path to the CSV file
