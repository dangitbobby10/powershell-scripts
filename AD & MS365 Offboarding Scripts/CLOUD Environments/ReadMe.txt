Requirements:
• PS 5.1
• Following PS Modules installed: MSOnline, ExchangeOnlineManagement, AzureAD


1: Obtain both scripts ('CLOUD -- Offboarding.ps1' & 'LicenseFriendlyNamesScript'). Notate 'LicenseFriendlyNamesScript's path as you will need to specify it in the 'CLOUD -- Offboarding.ps1' script as a defined variable.
2: Confirm you have the following powershell modules installed in PS 5.1:
	• MSOnline
	• ExchangeOnlineManagement
	• AzureAD
	
	(here's how to install them in case you dont have them:)
	Install-Module -Name MSOnline
	Install-Module -Name ExchangeOnlineManagement
	Install-Module -Name AzureAD

	import-module -name "repeat the same 3 modules above)