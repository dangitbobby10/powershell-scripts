# License SKU to friendly name mapping

<# 
————————————————————————
Retrieve SkuPartNumber:
————————————————————————
	
    
    Connect-MsolService
	Get-MsolAccountSku | Select-Object SkuPartNumber


Then match those up with this:
https://learn.microsoft.com/en-us/entra/identity/users/licensing-service-plan-reference

#>

$LicenseFriendlyNames = @{
	"AAD_PREMIUM" = "Azure Active Directory Premium P1"
	"AAD_PREMIUM_P2" = "Azure Active Directory Premium P2"
	"EMS" = "Enterprise Mobility + Security E3"
	"EXCHANGEARCHIVE_ADDON" = "Exchange Online Archiving for Exchange Online"
	"INTUNE_A" = "Intune"
	"MCOMEETADV" = "Microsoft 365 Audio Conferencing"
	"POWERAPPS_VIRAL" = "Microsoft Power Apps Plan 2 Trial"
	"FLOW_FREE" = "Microsoft Power Automate Free"
	"STREAM" = "Microsoft Stream Trial"
	"STANDARDPACK" = "Office 365 E1"
	"ENTERPRISEPACK" = "Office 365 E3"
	"ENTERPRISEPREMIUM" = "Office 365 E5"
	"POWER_BI_STANDARD" = "Power BI (free)"
	"POWER_BI_PRO" = "Power BI Pro"
	"PROJECTPROFESSIONAL" = "Project Plan 3"
	"VISIOONLINE_PLAN1" = "Visio Plan 1"
	"VISIOCLIENT" = "Visio Plan 2"
	"FORMS_PRO" = "Dynamics 365 Customer Voice Trial"
}