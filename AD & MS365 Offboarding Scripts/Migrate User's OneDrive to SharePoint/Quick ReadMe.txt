This script does work but I still consider it in a beta state as there are still a number of things I want to do to clean it up.

🔑 Key Prerequisits before running the script:
	• Powershell 7
	• PnP.Powershell Module (script installs if not detected -- "-scope CurrentUser")
	• Assumes you've configured an 'App Registration' for the PnPOnline module with Certificate-Based Authentication
		(If you haven't set this up I'll see what I can do to do a write up on the process but it's pretty easy and very Google-able)

#---------------------------------------------
# Azure App Reg Cert Creation (self-signed)
#---------------------------------------------
#create the cert
	$mycert = New-SelfSignedCertificate -DnsName "contoso.com" -CertStoreLocation "cert:\LocalMachine\My" -NotAfter (Get-Date).AddYears(4) -KeySpec KeyExchange -FriendlyName "PnPOnline automated scripts"

#view thumbprint of cert: 
	$mycert | Select-Object -Property Subject,Thumbprint,NotBefore,NotAfter

#export
	$mycert | Export-Certificate -FilePath "c:\mycertificates\pnponlineautomatedscripts.cer"

#and then the pfx export
	$mycert | Export-PfxCertificate -FilePath "c:\mycertificates\pnponlineautomatedscripts.pfx" -Password $(ConvertTo-SecureString -String "pass123word" -AsPlainText -Force)
#---------------------------------------------

■■■■■■■■■■■■■■■■■■■■
MSGraph App Reg API Permissions
■■■■■■■■■■■■■■■■■■■■
» Sites:
   • Site.ReadWrite.All
	
» Team Store:
   • TeamStore.ReadWrite.All

» User:
   • User.ReadWrite.All
■■■■■■■■■■■■■■■■■■■■

Again, I'm sorry for the lack of documentation as it is very late and I am very tired and hungry.

