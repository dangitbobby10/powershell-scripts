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

	import-module -name "repeat the same 3 modules above"
	
--------------------------------------------------------------------------------------------
	~~~~~~~~~~~~~~~~~~~~~~~~
	 Key Defined Variables:
	~~~~~~~~~~~~~~~~~~~~~~~~
	
	Just to be extra safe - if you comment out any of these key variables, be sure to find all the references in the script and comment/remove those as well to prevent errors.
	
	When you open the script in an editor, the key variables to define/comment out are:
	
		# Define the path to the "License Friendly Names Script" that transforms MS365 licenses from SKU to Friendly Names (e.g. "ENTERPRISEPACK" = "Office 365 E3")
		$LicenseFriendlyNamesScript = ""   # This is where you define where the 'LicenseFriendlyNamesScript's path

		# Define the 'Date' Variable for the CSV export file
		$date = Get-Date -Format "MM-dd-yyyy"	# I like it this way, but have at it

		# Define the path to the CSV file
		# (only change the value insde " ". Be sure to keep { } intact as it is used later as a script block if you want to keep $username and/or $date in your CSV files name.)
		$csvFilePath = { "c:\users\$env:username\desktop\Offboarding - $username $date.csv" }

--------------------------------------------------------------------------------------------
	~~~~~~~~~~~~~~
	 Script Flow:
	~~~~~~~~~~~~~~
	1: Connects to the x3 PS Modules (MSOnline, ExchangeOnline, AzureAD) -- will be prompted x3 times to log in.
	2: Prompts script executor for the following information:
		• Offboarded User's Full Email
		• Forwarding Address
		• Delegate 1
		• Delegate 2
		• Delegate 3
		• SendAs 1
		• SendAs 2
		• SendAs 3
		• Out of Office
		
	3: Pulls the following information an exports them to a CSV file and saves it on the script executor's desktop.
		• First Name
		• Last Name
		• Display Name
		• Email Address
		• UPN
		• Admin Roles
		• OnlineArchive Status
		• Litigation Hold Status
		• Job Title
		• Department
		• Mobile Phone
		• MS365 Groups
		• Forwarding To
		• Delegates
		• SendAs
		• SendOnBehalf
		• Licenses
		
	4: Force Block Signin for MS365
	5: Revokes MS365 & Azure Sessions
	6: Reset's AD password with a 21 random character generator	
	7: Removes from detected "MS365/Azure Admin Roles"	
	8: Hides account from the GAL
	9: Updates DisplayName to "Offboarded - $Displayname"	
	10: Converts Mailbox to Shared Mailbox and pauses for 2 minutes
	11: Removes All MS365 Licenses -- but first:
		• Analyzes Mailbox for size, Litigation Hold, and Online Archive.
			IF: Mailbox > 50 GB -- does not remove E3 & E5 licenses
			IF: LitHold is enabled -- does not remove E3 & E5 licenses (plan requires Exchange Online 2 so if you have a license that does that, slap in here)
			IF: Online Archiving Is Enabled -- does not remove E3, E5, & 'Exchange Online Archiving for Exchange' licenses	
	12: Sets Forwarding as defined in step 2
	13: Sets Delegates defined in step 2
	14: Sets SendAs defined in step 2
	15: Sets "Out of Office" defined in step 2
	16: Removes all MS365/Azure groups
	17: Prompts script executor if another user needs to be offboarded.
		• If yes, loops back to step 2
		• If no, ends script.
--------------------------------------------------------------------------------------------