How to use the EntraTermination.ps1 script

The script will check/install the required PowerShell modules you need to execute.

Step 1: Create a new App Registration in Azure and configure the API Permissions based on the “Graph API Permissions.PNG” file
	• App Registration > New Registration > Name = “TerminationScript” > Accounts in this org directory only

Step 2: Fill in the “Term_User.csv” 
	• Term_User_UPN = email address of user being off boarded
	• Delegate1-3 — email address of user needing delegate access
	• OOO — Out-of-Office response — leaving blank will not action this step	
	
	* You can off boarding multiple users at the same time using this method

Step 3: Execute Script, approve any interactive logins (there should be 2, one for Graph and one for ExchangeOnline)

Step 4: A CSV file will be exported with the users Group membership and licenses in the “\Exports\” folder

————————————————————————————————————————
Script Action:

1: Blocks Sign-in
2: Revoke Session
3: Hide from GAL
4: Converts to Shared Mailbox
5: Renames Displayname to “%Displayname% - Email Archive”
6: Exports Group Membership and Licenses to CSV in \Exports\ folder (does not remove groups or licenses)
7: Configures Delegates according to “Term_User.csv”
8: Configures OOO according to “Term_User.csv”
9: Disconnects from Graph and ExchangeOnline modules