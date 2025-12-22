Requirements:
• PS 5.1
• Following PS Modules installed: MSOnline, ExchangeOnlineManagement, AzureAD


1: Obtain both scripts ('HYBRID All OUs -- Offboarding.ps1' & 'LicenseFriendlyNamesScript'). Notate 'LicenseFriendlyNamesScript's path as you will need to specify it in the 'HYBRID All OUs -- Offboarding.ps1' script as a defined variable.
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
		
		# Define the domain controller
		$domainController = "" # I would suggest the DC that your AADConnect server points to

		# Define the (primary not staging) AADConnect Server
			$AADSyncServer = ""

		# Define the server FileServer
			$FileServer = ""        # This is part 1/3 of the variables to move a user's data to a secure folder on the FS.

		# Define the 'Offboarded Users' secure folder on the Folder server (ie typically lives in an IT or HR folder)
			$fs_offboardFolder = "" # Part 2/3 of moving a user's data to a secure folder

		# Define the User's HomeDirectory -- If you have homedirectories configured for your users in AD, the script will use that instead of this. No need to comment variable out.
			$manual_homedirectory = "" # what my copy/paste said - part 3/3 of moving a user's data to a secure folder on the FS.

		# Define "Disabled Users" OU
			$ou_path = ""           #"OU=Disabled Users,DC=contoso,DC=com"

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
	2: Imports AD module
	3: Prompts script executor for the following information:
		• Username (just the username, not email)
		• Forwarding Address
		• Delegate 1
		• Delegate 2
		• Delegate 3
		• SendAs 1
		• SendAs 2
		• SendAs 3
		• Out of Office
		
	4: Pulls the following information an exports them to a CSV file and saves it on the script executor's desktop.
		• AD Username
		• AD Description
		• AD Organizational Unit
		• AD Email
		• AD UPN
		• AD IPPhone
		• AD Groups
		• MS365 Groups
		• MS365 Admin Roles
		• Online Archive Status
		• Litigation Hold Status
		• Forwarding To
		• Delegates
		• SendAs
		• SendOnBehalf
		• Licenses
		
	5: Disables AD Account
	6: Reset's AD password with a 21 random character generator	
	7: If $fileserver, $fs_offboardFolder, AND $homedirectory/$manual_homedirectory has value, moves the user's data in the secure folder.
		7a: creates a job that will perform a check at the end of the script if the file transfer has completed	
	8: Updates AD Description to "Disabled on (current date)"
	9: Removes AD IPPhone field
	10: Removes ALL groups except for: Domain Users
	11: Hides account from the GAL
	12: Move AD User to "Disabled Users" OU
	13: Updates DisplayName to "Offboarded - $Displayname"
	14: Performs an AADConnect DeltaSync and pauses for 2 minutes
	15: Force Block Signin for MS365
	16: Revokes MS365 & Azure Sessions	
	17: Removes from detected "MS365/Azure Admin Roles"
	18: Converts Mailbox to Shared Mailbox and pauses for 2 minutes
	19: Removes All MS365 Licenses -- but first:
		• Analyzes Mailbox for size, Litigation Hold, and Online Archive.
			IF: Mailbox > 50 GB -- does not remove E3 & E5 licenses
			IF: LitHold is enabled -- does not remove E3 & E5 licenses (plan requires Exchange Online 2 so if you have a license that does that, slap in here)
			IF: Online Archiving Is Enabled -- does not remove E3, E5, & 'Exchange Online Archiving for Exchange' licenses	
	
	20: Sets Forwarding as defined in step 3
	
	21: Sets Delegates defined in step 3
	22: Sets SendAs defined in step 3
	23: Sets "Out of Office" defined in step 3
	24: Removes all MS365/Azure groups
	*25: If configured, checks if status if step 7a's transfer job.
	26: Prompts script executor if another user needs to be offboarded.
		• If yes, loops back to step 3
		• If no, ends script.
--------------------------------------------------------------------------------------------