# License SKU to friendly name mapping

<# 
Note from Robert talking about lines 242-259:
So I went a little overboard and grabbed the the 'groupassigninglicense' attributes as well but that was set up early in this script's infancy and I'm too lazy to 
take it out.

Really you only need the 'SkuPartNumber' -- but not the 'AccountSkuID.' My script, when it checks against the user's accountskuid, it strips the tenant part out of it and
sets the variable $sku which it then runs this script to grab its friendly name variables. 

Could I have adjusted my script to not strip and check? I could have, but then my tenant's ID is listed in this script and if you use it, your script.
Half lazy/half security? Who knows.





        Anyway, here's how you can grab your licenses:
            Connect-MsolService
            Get-MsolAccountSku | Select-Object SkuPartNumber

        Then match those up with this:
            https://learn.microsoft.com/en-us/entra/identity/users/licensing-service-plan-reference

#>

$LicenseFriendlyNames = @{
    "078d2b04-f1bd-4111-bbd4-b4b1b354cef4" = "Azure Active Directory Premium P1"
	"AAD_PREMIUM" = "Azure Active Directory Premium P1"
	"84a661c4-e949-4bd2-a560-ed7766fcaf2b" = "Azure Active Directory Premium P2"
	"AAD_PREMIUM_P2" = "Azure Active Directory Premium P2"
    "efccb6f7-5641-4e0e-bd10-b4976e1bf68e" = "Enterprise Mobility + Security E3"
    "EMS" = "Enterprise Mobility + Security E3"
    "ee02fd1b-340e-4a4b-b355-4a514e4c8943" = "Exchange Online Archiving for Exchange Online"
	"EXCHANGEARCHIVE_ADDON" = "Exchange Online Archiving for Exchange Online"
    "061f9ace-7d42-4136-88ac-31dc755f143f" = "Intune"
	"INTUNE_A" = "Intune"
    "0c266dff-15dd-4b49-8397-2bb16070ed52" = "Microsoft 365 Audio Conferencing"
	"MCOMEETADV" = "Microsoft 365 Audio Conferencing"
    "dcb1a3ae-b33f-4487-846a-a640262fadf4" = "Microsoft Power Apps Plan 2 Trial"
	"POWERAPPS_VIRAL" = "Microsoft Power Apps Plan 2 Trial"
    "f30db892-07e9-47e9-837c-80727f46fd3d" = "Microsoft Power Automate Free"
	"FLOW_FREE" = "Microsoft Power Automate Free"
    "1f2f344a-700d-42c9-9427-5cea1d5d7ba6" = "Microsoft Stream Trial"
	"STREAM" = "Microsoft Stream Trial"
    "18181a46-0d4e-45cd-891e-60aabd171b4e" = "Office 365 E1"
	"STANDARDPACK" = "Office 365 E1"
    "6fd2c87f-b296-42f0-b197-1e91e994b900" = "Office 365 E3"
	"ENTERPRISEPACK" = "Office 365 E3"
    "c7df2760-2c81-4ef7-b578-5b5392b571df" = "Office 365 E5"
	"ENTERPRISEPREMIUM" = "Office 365 E5"
    "a403ebcc-fae0-4ca2-8c8c-7a907fd6c235" = "Power BI (free)"
	"POWER_BI_STANDARD" = "Power BI (free)"
    "f8a1db68-be16-40ed-86d5-cb42ce701560" = "Power BI Pro"
	"POWER_BI_PRO" = "Power BI Pro"
    "53818b1b-4a27-454b-8896-0dba576410e6" = "Project Plan 3"
	"PROJECTPROFESSIONAL" = "Project Plan 3"
    "4b244418-9658-4451-a2b8-b5e2b364e9bd" = "Visio Plan 1"
	"VISIOONLINE_PLAN1" = "Visio Plan 1"
    "c5928f49-12ba-48f7-ada3-0d743a3601d5" = "Visio Plan 2"
	"VISIOCLIENT" = "Visio Plan 2"
	"bc946dac-7877-4271-b2f7-99d2db13cd2c" = "Dynamics 365 Customer Voice Trial"
	"FORMS_PRO" = "Dynamics 365 Customer Voice Trial"
}