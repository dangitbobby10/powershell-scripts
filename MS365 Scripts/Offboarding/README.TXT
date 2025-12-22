How to use the EntraTermination.ps1 script

The script will check/install the required PowerShell modules you need to execute. This script is using the INERACTIVE login method.
This script will also ONLY offboard users that are Full Cloud - if you have DirSync, convert the account into Cloud then run the script.

Step 1: Fill in the “Term_User.csv” 
	• Term_User_UPN = email address of user being off boarded
	• Delegate1-3 — email address of user needing delegate access
	• OOO — Out-of-Office response — leaving blank will not action this step	
	
	* You can off boarding multiple users at the same time using this method

Step 2: Execute Script, approve any interactive logins (there should be 2, one for Graph and one for ExchangeOnline)

Step 3: A CSV file will be exported with the users Group membership and licenses in the “\Exports\” folder

* If you run into any errors, review the \Logs\TerminationErrors.csv file.

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