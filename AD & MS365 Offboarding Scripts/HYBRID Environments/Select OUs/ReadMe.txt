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
