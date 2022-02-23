#Requires -Version 3.0

<#
.SYNOPSIS
	Creates a basic Health Check of a Citrix PVS 5.x or later farm.
.DESCRIPTION
    Creates a basic health check of a Citrix PVS 5.x or later farm in plain text, 
    HTML, Word, or PDF.

    This script has been tested with PVS 2112.

    Creates a document named after the PVS farm.

    The script must run from an elevated PowerShell session.

    NOTE: The account used to run this script must have at least Read access to the SQL 
    Server that holds the Citrix Provisioning databases.

    This script is written using the old string-based crappy PowerShell because it 
    supports PVS 5.x.

    For the email stuff to work, this script requires at least PowerShell V3.
	
.PARAMETER AdminAddress
	Specifies the name of a PVS server that the PowerShell script connects to. 
	Using this parameter requires the script to run from an elevated PowerShell session.
	Starting with V1.11 of the script, this requirement is now checked.

	This parameter has an alias of AA
.PARAMETER Domain
	Specifies the domain used for the AdminAddress connection. 

	Default value is contained in $env:UserDomain
.PARAMETER User
	Specifies the user used for the AdminAddress connection. 

	Default value is contained in $env:username
.PARAMETER CSV
	Creates a CSV file for each Appendix.
	The default value is False.
	
	Output CSV filename is in the format:
	
	PVSFarmName_HealthCheck_Appendix#_NameOfAppendix.csv
	
	For example:
		TNPVSFarm_HealthCheck_AppendixA_AdvancedServerItems1.csv
		TNPVSFarm_HealthCheck_AppendixB_AdvancedServerItems2.csv
		TNPVSFarm_HealthCheck_AppendixC_ConfigWizardItems.csv
		TNPVSFarm_HealthCheck_AppendixD_ServerBootstrapItems.csv
		TNPVSFarm_HealthCheck_AppendixE_DisableTaskOffloadSetting.csv	
		TNPVSFarm_HealthCheck_AppendixF_PVSServices.csv
		TNPVSFarm_HealthCheck_AppendixG_vDiskstoMerge.csv	
		TNPVSFarm_HealthCheck_AppendixH_EmptyDeviceCollections.csv	
		TNPVSFarm_HealthCheck_AppendixI_UnassociatedvDisks.csv	
		TNPVSFarm_HealthCheck_AppendixJ_BadStreamingIPAddresses.csv	
		TNPVSFarm_HealthCheck_AppendixK_MiscRegistryItems.csv
		TNPVSFarm_HealthCheck_AppendixL_vDisksConfiguredforServerSideCaching.csv	
		TNPVSFarm_HealthCheck_AppendixM_MicrosoftHotfixesandUpdates.csv
		TNPVSFarm_HealthCheck_AppendixN_InstalledRolesandFeatures.csv
		TNPVSFarm_HealthCheck_AppendixO_PVSProcesses.csv
		TNPVSFarm_HealthCheck_AppendixP_ItemsToReview.csv
		TNPVSFarm_HealthCheck_AppendixQ_ServerComputerItemsToReview.csv
		TNPVSFarm_HealthCheck_AppendixQ_ServerDriveItemsToReview.csv
		TNPVSFarm_HealthCheck_AppendixQ_ServerProcessorItemsToReview.csv
		TNPVSFarm_HealthCheck_AppendixQ_ServerNICItemsToReview.csv
		TNPVSFarm_HealthCheck_AppendixR_CitrixInstalledComponents.csv
.PARAMETER Dev
	Clears errors at the beginning of the script.
	Outputs all errors to a text file at the end of the script.
	
	This is used when the script developer requests more troubleshooting data.
	The text file is placed in the same folder from where the script is run.
	
	This parameter is disabled by default.
.PARAMETER Folder
	Specifies the optional output folder to save the output report. 
.PARAMETER Log
	Generates a log file for troubleshooting.
.PARAMETER ScriptInfo
	Outputs information about the script to a text file.
	The text file is placed in the same folder from where the script is run.
	
	This parameter is disabled by default.
	This parameter has an alias of SI.
.PARAMETER SmtpServer
	Specifies the optional email server to send the output report. 
.PARAMETER SmtpPort
	Specifies the SMTP port. 
	The default port is 25.
.PARAMETER UseSSL
	Specifies whether to use SSL for the SmtpServer.
	THe default is False.
.PARAMETER From
	Specifies the username for the From email address.
	If SmtpServer is used, this is a required parameter.
.PARAMETER To
	Specifies the username for the To email address.
	If SmtpServer is used, this is a required parameter.
.EXAMPLE
	PS C:\PSScript > .\PVS_HealthCheck_V2.ps1
	
	Uses all Default values.
	LocalHost for AdminAddress.
.EXAMPLE
	PS C:\PSScript > .\PVS_HealthCheck_V2.ps1 -AdminAddress PVS1 -User cwebster -Domain WebstersLab

	Use this example to run the script against a PVS Farm in another domain or forest.
	
	Will use:
		PVS1 for AdminAddress.
		cwebster for User.
		WebstersLab for Domain.

	Uses Get-Credential to prompt for the password.
.EXAMPLE
	PS C:\PSScript > .\PVS_HealthCheck_V2.ps1 -AdminAddress PVS1 -User cwebster

	Will use:
		PVS1 for AdminAddress.
		cwebster for User.
		$env:UserDnsDomain for the Domain.

	Uses Get-Credential to prompt for the password.
.EXAMPLE
	PS C:\PSScript > .\PVS_HealthCheck_V2.ps1 -Folder \\FileServer\ShareName
	
	Output file will be saved in the path \\FileServer\ShareName
.EXAMPLE
	PS C:\PSScript > .\PVS_HealthCheck_V2.ps1 -SmtpServer mail.domain.tld -From 
	XDAdmin@domain.tld -To ITGroup@domain.tld -ComputerName DHCPServer01
	
	Script uses the email server mail.domain.tld, sending from XDAdmin@domain.tld and 
	sending to ITGroup@domain.tld.
	
	If the current user's credentials are not valid to send an email, the script prompts 
	the user to enter valid credentials.
.EXAMPLE
	PS C:\PSScript > .\PVS_HealthCheck_V2.ps1 -Dev -ScriptInfo -Log
	
	Creates a text file named PVSHealthCheckScriptErrors_yyyy-MM-dd_HHmm.txt that 
	contains up to the last 250 errors reported by the script.
	
	Creates a text file named PVSHealthCheckScriptInfo_yyyy-MM-dd_HHmm.txt that 
	contains all the script parameters and other basic information.
	
	Creates a text file for transcript logging named 
	PVSHealthCheckScriptTranscript_yyyy-MM-dd_HHmm.txt.
.EXAMPLE
	PS C:\PSScript > .\PVS_HealthCheck_V2.ps1 -CSV
	
	Uses all Default values.
	LocalHost for AdminAddress.
	Creates a CSV file for each Appendix.
.EXAMPLE
	PS C:\PSScript > .\PVS_HealthCheck_V2.ps1 -SmtpServer mail.domain.tld -From 
	XDAdmin@domain.tld -To ITGroup@domain.tld	

	The script uses the email server mail.domain.tld, sending from XDAdmin@domain.tld 
	and sending to ITGroup@domain.tld.

	The script uses the default SMTP port 25 and does not use SSL.

    If the current user's credentials are not valid to send an email, the script prompts 
    the user to enter valid credentials.
.EXAMPLE
	PS C:\PSScript > .\PVS_HealthCheck_V2.ps1 -SmtpServer mailrelay.domain.tld	-From 
	Anonymous@domain.tld -To ITGroup@domain.tld	

	***SENDING UNAUTHENTICATED EMAIL***

	The script uses the email server mailrelay.domain.tld, sending from anonymous@domain.tld and sending to ITGroup@domain.tld.

	To send an unauthenticated email using an email relay server requires the From email 
	account to use the name Anonymous.

	The script uses the default SMTP port 25 and does not use SSL.
	
	***GMAIL/G SUITE SMTP RELAY***
	https://support.google.com/a/answer/2956491?hl=en
	https://support.google.com/a/answer/176600?hl=en

	To send an email using a Gmail or g-suite account, you may have to turn ON the "Less 
	secure app access" option on your account.
	***GMAIL/G SUITE SMTP RELAY***

	The script generates an anonymous, secure password for the anonymous@domain.tld 
	account.
.EXAMPLE
	PS C:\PSScript > .\PVS_HealthCheck_V2.ps1 -SmtpServer 
	labaddomain-com.mail.protection.outlook.com -UseSSL -From 
	SomeEmailAddress@labaddomain.com -To ITGroupDL@labaddomain.com	

	***OFFICE 365 Example***

	https://docs.microsoft.com/en-us/exchange/mail-flow-best-practices/how-to-set-up-a-multifunction-device-or-application-to-send-email-using-office-3
	
	This uses Option 2 from the above link.
	
	***OFFICE 365 Example***

	The script uses the email server labaddomain-com.mail.protection.outlook.com, sending f
	rom SomeEmailAddress@labaddomain.com and sending to ITGroupDL@labaddomain.com.

	The script uses the default SMTP port 25 and SSL.
.EXAMPLE
	PS C:\PSScript > .\PVS_HealthCheck_V2.ps1 -SmtpServer smtp.office365.com -SmtpPort 587
	-UseSSL -From Webster@CarlWebster.com -To ITGroup@CarlWebster.com	

	The script uses the email server smtp.office365.com on port 587 using SSL, sending from 
	webster@carlwebster.com and sending to ITGroup@carlwebster.com.

    If the current user's credentials are not valid to send an email, the script prompts 
    the user to enter valid credentials.
.EXAMPLE
	PS C:\PSScript > .\PVS_HealthCheck_V2.ps1 -SmtpServer smtp.gmail.com -SmtpPort 587 -UseSSL 
	-From Webster@CarlWebster.com -To ITGroup@CarlWebster.com	

	*** NOTE ***
	To send an email using a Gmail or g-suite account, you may have to turn ON the "Less 
	secure app access" option on your account.
	*** NOTE ***
	
	The script uses the email server smtp.gmail.com on port 587 using SSL, sending from 
	webster@gmail.com and sending to ITGroup@carlwebster.com.

    If the current user's credentials are not valid to send an email, the script prompts 
    the user to enter valid credentials.
.INPUTS
	None.  You cannot pipe objects to this script.
.OUTPUTS
	No objects are output from this script.  This script creates a text file and optional 
	CSV files.
.NOTES
	NAME: PVS_HealthCheck_V2.ps1
	VERSION: 2.00
	AUTHOR: Carl Webster (with much help from BG a, now former, Citrix dev)
	LASTEDIT: February 23, 2022
#>


#thanks to @jeffwouters for helping me with these parameters
[CmdletBinding(SupportsShouldProcess = $False, ConfirmImpact = "None", DefaultParameterSetName = "") ]

Param(
	[parameter(Mandatory=$False)] 
	[Alias("AA")]
	[string]$AdminAddress="",

	[parameter(Mandatory=$False)] 
	[string]$Domain=$env:UserDomain,

	[parameter(Mandatory=$False)] 
	[string]$User=$env:username,

	[parameter(Mandatory=$False)] 
	[switch]$CSV=$False,

	[parameter(Mandatory=$False)] 
	[Switch]$Dev=$False,
	
	[parameter(Mandatory=$False)] 
	[string]$Folder="",

	[parameter(Mandatory=$False)] 
	[Switch]$Log=$False,
	
	[parameter(Mandatory=$False)] 
	[Alias("SI")]
	[Switch]$ScriptInfo=$False,
	
	[parameter(Mandatory=$False)] 
	[string]$SmtpServer="",

	[parameter(Mandatory=$False)] 
	[int]$SmtpPort=25,

	[parameter(Mandatory=$False)] 
	[switch]$UseSSL=$False,

	[parameter(Mandatory=$False)] 
	[string]$From="",

	[parameter(Mandatory=$False)] 
	[string]$To=""
	
	)


#Carl Webster, CTP
#webster@carlwebster.com
#@carlwebster on Twitter
#http://www.CarlWebster.com
#V2 script created February 23, 2022
#released to the community on 
#V2.00 is based on 1.24
#


#region PoSH prereqs
Function CheckOnPoSHPrereqs
{
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Checking for McliPSSnapin"
	If(!(Check-NeededPSSnapins "McliPSSnapIn"))
	{
		#We're missing Citrix Snapins that we need
		#changed in 1.23 to the console installation path
		#this should return <DriveLetter:>\Program Files\Citrix\Provisioning Services Console\
		$PFiles = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Citrix\ProvisioningServices' -Name ConsoleTargetDir -ErrorAction SilentlyContinue)|Select-Object -ExpandProperty ConsoleTargetDir
		$PVSDLLPath = Join-Path -Path $PFiles -ChildPath "McliPSSnapIn.dll"
		#Let's see if the DLLs can be registered
		If(Test-Path $PVSDLLPath -EA 0)
		{
			Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Searching for the 32-bit .Net V2 snapin"
			$installutil = $env:systemroot + '\Microsoft.NET\Framework\v2.0.50727\installutil.exe'
			If(Test-Path $installutil -EA 0)
			{
				Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): `tAttempting to register the 32-bit .Net V2 snapin"
				&$installutil $PVSDLLPath > $Null
			
				If(!$?)
				{
					Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): `t`tUnable to register the 32-bit V2 PowerShell Snap-in."
				}
				Else
				{
					Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): `t`tRegistered the 32-bit V2 PowerShell Snap-in."
				}
			}
			Else
			{
				Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): `tNo 32-bit .Net V2 snapin found"
			}
	
			Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Searching for the 64-bit .Net V2 snapin"
			$installutil = $env:systemroot + '\Microsoft.NET\Framework64\v2.0.50727\installutil.exe'
			If(Test-Path $installutil -EA 0)
			{
				Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): `tAttempting to register the 64-bit .Net V2 snapin"
				&$installutil $PVSDLLPath > $Null
			
				If(!$?)
				{
					Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): `t`tUnable to register the 64-bit V2 PowerShell Snap-in."
				}
				Else
				{
					Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): `t`tRegistered the 64-bit V2 PowerShell Snap-in."
				}
			}
			Else
			{
				Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): `tNo 64-bit .Net V2 snapin found"
			}
		}
		Else
		{
			Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Unable to find $PVSDLLPath"
		}
		
		If(Test-Path $PVSDLLPath -EA 0)
		{
			Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Searching for the 32-bit .Net V4 snapin"
			$installutil = $env:systemroot + '\Microsoft.NET\Framework\v4.0.30319\installutil.exe'
			If(Test-Path $installutil -EA 0)
			{
				Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): `tAttempting to register the 32-bit .Net V4 snapin"
				&$installutil $PVSDLLPath > $Null
			
				If(!$?)
				{
					Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): `t`tUnable to register the 32-bit V4 PowerShell Snap-in."
				}
				Else
				{
					Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): `t`tRegistered the 32-bit V4 PowerShell Snap-in."
				}
			}
			Else
			{
				Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): `tNo 32-bit .Net V4 snapin found"
			}
	
			Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Searching for the 64-bit .Net V4 snapin"
			$installutil = $env:systemroot + '\Microsoft.NET\Framework64\v4.0.30319\installutil.exe'
			If(Test-Path $installutil -EA 0)
			{
				Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): `tAttempting to register the 64-bit .Net V4 snapin"
				&$installutil $PVSDLLPath > $Null
			
				If(!$?)
				{
					Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): `t`tUnable to register the 64-bit V4 PowerShell Snap-in."
				}
				Else
				{
					Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): `t`tRegistered the 64-bit V4 PowerShell Snap-in."
				}
			}
			Else
			{
				Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): `tNo 64-bit .Net V4 snapin found"
			}
		}
		Else
		{
			Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Unable to find $PVSDLLPath"
		}
	
		Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Rechecking for McliPSSnapin"
		If(!(Check-NeededPSSnapins "McliPSSnapIn"))
		{
			#We're missing Citrix Snapins that we need
			Write-Error "
			`n`n
			`t`t
			Missing Citrix PowerShell Snap-ins Detected, check the console above for more information.
			`n`n
			`t`t
			Script will now close.
			`n`n
			"
			Exit
		}
		Else
		{
			Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Citrix PowerShell Snap-ins detected at $PVSDLLPath"
		}
	}
	Else
	{
		Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Citrix PowerShell Snap-ins detected."
	}

}
#endregion

#region remoting function
Function SetupRemoting
{
	#setup remoting if $AdminAddress is not empty
	[bool]$Script:Remoting = $False
	If(![System.String]::IsNullOrEmpty($AdminAddress))
	{
		#V1.23 changed to get-credentials with tip and code from Frank Lindenblatt of the PoSH Users Group Hanover (Germany)
		#This way the Password is not exposed in plaintext

		$credential = Get-Credential -Message "Enter the credentials to connect to $AdminAddress" -UserName "$Domain\$User"

		If($Null -ne $credential)
		{
			$netCred = $credential.GetNetworkCredential()
	
			$Domain   = "$($netCred.Domain)"
			$User     = "$($netCred.UserName)"
			$Password = "$($netCred.Password)"

			$error.Clear()
			mcli-run SetupConnection -p server="$($AdminAddress)",user="$($User)",domain="$($Domain)",password="$($Password)"

			If($error.Count -eq 0)
			{
				$Script:Remoting = $True
				Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): This script is being run remotely against server $($AdminAddress)"
				If(![System.String]::IsNullOrEmpty($User))
				{
					Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): User=$($User)"
					Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Domain=$($Domain)"
				}
			}
			Else 
			{
				Write-Warning "Remoting could not be setup to server $($AdminAddress)"
				$tmp = $Error[0]
				Write-Warning "Error returned is $tmp"
				Write-Warning "Script cannot continue"
				Exit
			}
		}
		Else 
		{
			Write-Warning "Remoting could not be setup to server $($AdminAddress)"
			Write-Warning "Credentials are invalid"
			Write-Warning "Script cannot continue"
			Exit
		}
	}
	Else
	{
		#added V1.17
		#if $AdminAddress is "", get actual server name
		If($AdminAddress -eq "")
		{
			$Script:AdminAddress = $env:ComputerName
		}
	}
}
#endregion

#region verify PVS services
Function VerifyPVSServices
{
	If($AdminAddress -eq "")
	{
		$tmp = $env:ComputerName
		Write-Verbose "$(Get-Date -Format G): Server name changed from localhost to $tmp"
	}
	Else
	{
		$tmp = $AdminAddress
	}
	
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Verifying PVS SOAP and Stream Services are running on $tmp"

	$soapserver = $Null
	$StreamService = $Null

	If($Script:Remoting)
	{
		$soapserver = Get-Service -ComputerName $AdminAddress -EA 0 | Where-Object {$_.DisplayName -like "*Citrix PVS Soap Server*"}
		$StreamService = Get-Service -ComputerName $AdminAddress -EA 0 | Where-Object {$_.DisplayName -like "*Citrix PVS Stream Service*"}
	}
	Else
	{
		$soapserver = Get-Service -EA 0 | Where-Object {$_.DisplayName -like "*Citrix PVS Soap Server*"}
		$StreamService = Get-Service -EA 0 | Where-Object {$_.DisplayName -like "*Citrix PVS Stream Service*"}
	}

	If($Null -eq $soapserver)
	{
		Write-Error "
		`n`n
		`t`t
		The Citrix PVS Soap Server service status on $tmp could not be determined.
		`n`n
		`t`t
		Script cannot continue.
		`n`n
		"
		Exit
	}
	Else
	{
		If($soapserver.Status -ne "Running")
		{
			$txt = "The Citrix PVS Soap Server service is not Started on server $tmp"
			Write-Error "
			`n`n
			`t`t
			$txt
			`n`n
			`t`t
			Script cannot continue.
			`n`n
			"
			Exit
		}
	}

	If($Null -eq $StreamService)
	{
		Write-Error "
		`n`n
		`t`t
		The Citrix PVS Stream Service service status on $tmp could not be determined.
		`n`n
		`t`t
		Script cannot continue.
		`n`n
		"
		Exit
	}
	Else
	{
		If($StreamService.Status -ne "Running")
		{
			$txt = "The Citrix PVS Stream Service service is not Started on server $tmp"
			Write-Error "
			`n`n
			`t`t
			$txt
			`n`n
			`t`t
			Script cannot continue.
			`n`n
			"
			Exit
		}
	}
}
#endregion

#region getpvsversion
Function GetPVSVersion
{
	#get PVS major version
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Getting PVS version info"

	$error.Clear()
	$tempversion = mcli-info version
	If($? -and $error.Count -eq 0)
	{
		#build PVS version values
		$version = new-object System.Object 
		ForEach($record in $tempversion)
		{
			$index = $record.IndexOf(':')
			If($index -gt 0)
			{
				$property = $record.SubString(0, $index)
				$value = $record.SubString($index + 2)
				Add-Member -inputObject $version -MemberType NoteProperty -Name $property -Value $value
			}
		}
	} 
	Else 
	{
		Write-Warning "PVS version information could not be retrieved"
		[int]$NumErrors = $Error.Count
		For($x=0; $x -le $NumErrors; $x++)
		{
			Write-Warning "Error(s) returned: " $error[$x]
		}
		Write-Error "
		`n`n
		`t`t
		Script is terminating
		`n`n
		"
		#without version info, script should not proceed
		Exit
	}

	$Script:PVSVersion              = $Version.mapiVersion.SubString(0,1)
	[version]$Script:PVSFullVersion = $Version.mapiVersion
}
#endregion

#region get PVS Farm functions
Function GetPVSFarm
{
	#build PVS farm values
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Build PVS farm values"
	#there can only be one farm
	$GetWhat = "Farm"
	$GetParam = ""
	$ErrorTxt = "PVS Farm information"
	$Script:Farm = BuildPVSObject $GetWhat $GetParam $ErrorTxt

	If($Null -eq $Script:Farm)
	{
		#without farm info, script should not proceed
		Write-Error "
		`n`n
		`t`t
		PVS Farm information could not be retrieved.
		`n`n
		`t`t
		Script is terminating.
		`n`n
		"
		Exit
	}

	[string]$Script:Title = "PVS Health Check Report for Farm $($Script:farm.FarmName)"
	SetFileName1 "$($Script:farm.FarmName)_HealthCheck" #V1.16 add _Assessment
}

Function SetFileName1
{
	Param([string]$OutputFileName)
	[string]$Script:FileName1 = "$($Script:pwdpath)\$($OutputFileName).txt"
}
#endregion

#region show script options
Function ShowScriptOptions
{
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): "
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): "
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): AdminAddress       : $($AdminAddress)"
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): CSV                : $($CSV)"
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Dev                : $($Dev)"
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Domain             : $($Domain)"
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Filename1          : $($Script:filename1)"
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Folder             : $($Script:pwdpath)"
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): From               : $($From)"
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Log                : $($Log)"
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): PVS Farm Name      : $($Script:farm.farmName)"
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): PVS Version        : $($Script:PVSFullVersion)"
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): ScriptInfo         : $($ScriptInfo)"
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Smtp Port          : $($SmtpPort)"
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Smtp Server        : $($SmtpServer)"
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Title              : $($Script:Title)"
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): To                 : $($To)"
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Use SSL            : $($UseSSL)"
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): User               : $($User)"
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): "
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): OS Detected        : $($Script:RunningOS)"
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): PoSH version       : $($Host.Version)"
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): PSCulture          : $($PSCulture)"
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): PSUICulture        : $($PSUICulture)"
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): "
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Script start       : $($Script:StartTime)"
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): "
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): "
}
#endregion

#region process pvs farm functions
Function Get-IPAddress
{
	#V1.16 added new function
	Param([string]$ComputerName)
	
	If( ! [string]::ISNullOrEmpty( $computername ) )
	{
		$IPAddress = "Unable to determine"
		
		Try
		{
			$IP = Test-Connection -ComputerName $ComputerName -Count 1 | Select-Object IPV4Address
		}
		
		Catch
		{
			$IP = "Unable to resolve IP address"
		}

		If($? -and $Null -ne $IP -and $IP -ne "Unable to resolve IP address")
		{
			$IPAddress = $IP.IPV4Address.IPAddressToString
		}
	}
	Else
	{
		$IPAddress = ""
	}
	
	Return $IPAddress
}

Function ProcessPVSFarm
{
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Processing PVS Farm Information"

	$LicenseServerIPAddress = Get-IPAddress $Script:farm.licenseServer #added in V1.16
	
	#V1.17 see if the database server names contain an instance name. If so, remove it
	#V1.18 add test for port number - bug found by Johan Parlevliet 
	#V1.18 see if the database server names contain a port number. If so, remove it
	#V1.18 optimized code supplied by MBS
	$dbServer = $Script:farm.databaseServerName
	If( ( $inx = $dbServer.IndexOfAny( ',\' ) ) -ge 0 )
	{
		#strip the instance name and/or port name, if present
		Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Removing '$( $dbServer.SubString( $inx ) )' from SQL server name to get IP address"
		$dbServer = $dbServer.SubString( 0, $inx )
		Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): dbServer now '$dbServer'"
	}
	$SQLServerIPAddress = Get-IPAddress $dbServer #added in V1.16
	
	$dbServer = $Script:farm.failoverPartnerServerName
	If( ( $inx = $dbServer.IndexOfAny( ',\' ) ) -ge 0 )
	{
		#strip the instance name and/or port name, if present
		Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Removing '$( $dbServer.SubString( $inx ) )' from SQL server name to get IP address"
		$dbServer = $dbServer.SubString( 0, $inx )
		Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): dbServer now '$dbServer'"
	}
	$FailoverSQLServerIPAddress = Get-IPAddress $dbServer #added in V1.16
	
	#general tab
	Line 0 "PVS Farm Name`t`t`t: " $Script:farm.farmName
	Line 0 "Version`t`t`t`t: " $Script:PVSFullVersion
	
	#security tab
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): `tProcessing Security Tab"
	Line 0 "Security"
	Line 1 "Groups with Farm Administrator access:"
	#build security tab values
	$GetWhat = "authgroup"
	$GetParam = "farm = 1"
	$ErrorTxt = "Groups with Farm Administrator access"
	$authgroups = BuildPVSObject $GetWhat $GetParam $ErrorTxt

	If($Null -ne $AuthGroups)
	{
		ForEach($Group in $authgroups)
		{
			If($Group.authGroupName)
			{
				Line 2 $Group.authGroupName
			}
		}
	}

	#groups tab
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): `tProcessing Groups Tab"
	Line 0 "Groups"
	Line 1 "All the Security Groups that can be assigned access rights:"
	$GetWhat = "authgroup"
	$GetParam = ""
	$ErrorTxt = "Security Groups information"
	$authgroups = BuildPVSObject $GetWhat $GetParam $ErrorTxt

	If($Null -ne $AuthGroups)
	{
		ForEach($Group in $authgroups)
		{
			If($Group.authGroupName)
			{
				Line 2 $Group.authGroupName
			}
		}
	}

	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): `tProcessing Licensing Tab"
	Line 0 "License server name`t`t: " $Script:farm.licenseServer
	Line 0 "License server IP`t`t: " $LicenseServerIPAddress
	Line 0 "License server port`t`t: " $Script:farm.licenseServerPort
	If($Script:PVSVersion -eq "5")
	{
		Line 0 "Use Datacenter licenses for desktops if no Desktop licenses are available: " -nonewline
		If($farm.licenseTradeUp -eq "1")
		{
			Line 0 "Yes"
		}
		Else
		{
			Line 0 "No"
		}
	}

	If($Script:PVSFullVersion -ge "7.19")
	{
		Line 0 "Citrix Provisioning license type" ""
		If($farm.LicenseSKU -eq 0)  #fix in 1.21 uint LicenseSKU: LicenseSKU. 0 for on-premises, 1 for cloud. Min=0, Max=1, Default=0
		{
			Line 1 "On-Premises`t: " "Yes"
			Line 2 "Use Datacenter licenses for desktops if no Desktop licenses are available: " -nonewline
			If($farm.licenseTradeUp -eq "1")
			{
				Line 0 "Yes"
			}
			Else
			{
				Line 0 "No"
			}
			Line 1 "Cloud`t`t: " "No"
		}
		ElseIf($farm.LicenseSKU -eq 1)
		{
			Line 1 "On-Premises`t: " "No"
			Line 2 "Use Datacenter licenses for desktops if no Desktop licenses are available: No"
			Line 1 "Cloud`t`t: " "Yes"
		}
		Else
		{
			Line 1 "On-Premises: " "ERROR: Unable to determine the PVS License SKU Tpe"
		}
	}
	ElseIf($Script:PVSFullVersion -ge "7.13")
	{
		#	Fixed in the 1.23 the missing $DatacenterLicense variable (found by SHurjuk)

		If($farm.licenseTradeUp -eq "1" -or $farm.licenseTradeUp -eq $True)
		{
			$DatacenterLicense = "Yes"
		}
		Else
		{
			$DatacenterLicense = "No"
		}
		Line 1 "Use Datacenter licenses for desktops if no Desktop licenses are available: " $DatacenterLicense
	}

	Line 0 "Enable auto-add`t`t`t: " -nonewline
	If($farm.autoAddEnabled -eq "1")
	{
		Line 0 "Yes"
		Line 0 "Add new devices to this site`t: " $farm.DefaultSiteName
		$Script:FarmAutoAddEnabled = $True
	}
	Else
	{
		Line 0 "No"	
		$Script:FarmAutoAddEnabled = $False
	}	
	
	#options tab
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): `tProcessing Options Tab"
	Line 0 "Enable auditing`t`t`t: " -nonewline
	If($Script:farm.auditingEnabled -eq "1")
	{
		Line 0 "Yes"
	}
	Else
	{
		Line 0 "No"
		$obj1 = [PSCustomObject] @{
			ItemText = "Auditing is not enabled"
		}
		$null = $Script:ItemsToReview.Add($obj1)
	}
	Line 0 "Enable offline database support`t: " -nonewline
	If($Script:farm.offlineDatabaseSupportEnabled -eq "1")
	{
		Line 0 "Yes"	
	}
	Else
	{
		Line 0 "No"
		$obj1 = [PSCustomObject] @{
			ItemText = "Offline database support is not enabled"
		}
		$null = $Script:ItemsToReview.Add($obj1)
	}
	
	#7.11 Problem Report tab
	If($Script:PVSFullVersion -ge "7.11")
	{
		$GetWhat = "cisdata"
		$GetParam = ""
		$ErrorTxt = "Problem Report information"
		$ProblemReports = BuildPVSObject $GetWhat $GetParam $ErrorTxt
		
		If($Null -ne $ProblemReports)
		{
			If($ProblemReports.UserName -ne "")
			{
				Line 0 "Problem report Citrix Username`t: " $ProblemReports.UserName
				$obj1 = [PSCustomObject] @{
					ItemText = "Problem report Citrix Username is $($ProblemReports.UserName)"
				}
				$null = $Script:ItemsToReview.Add($obj1)
			}
		}
	}
	
	If($Script:PVSVersion -eq "6" -or $Script:PVSVersion -eq "7")
	{
		#vDisk Version tab
		Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): `tProcessing vDisk Version Tab"
		Line 0 "vDisk Version"
		Line 1 "Alert if number of versions from base image exceeds`t: " $Script:farm.maxVersions
		Line 1 "Default access mode for new merge versions`t`t: " -nonewline
		Switch ($Script:farm.mergeMode)
		{
			0   {Line 0 "Production"; Break }
			1   {Line 0 "Test"; Break }
			2   {Line 0 "Maintenance"; Break}
			Default {Line 0 "Default access mode could not be determined: $($Script:farm.mergeMode)"; Break}
		}
	}
	
	#status tab
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): `tProcessing Status Tab"
	Line 0 "Database server`t`t`t: " $Script:farm.databaseServerName
	Line 0 "Database server IP`t`t: " $SQLServerIPAddress
	Line 0 "Database instance`t`t: " $Script:farm.databaseInstanceName
	Line 0 "Database`t`t`t: " $Script:farm.databaseName
	Line 0 "Failover Partner Server`t`t: " $Script:farm.failoverPartnerServerName
	Line 0 "Failover Partner Server IP`t: " $FailoverSQLServerIPAddress
	Line 0 "Failover Partner Instance`t: " $Script:farm.failoverPartnerInstanceName
	If($Script:farm.adGroupsEnabled -eq "1")
	{
		Line 0 "Active Directory groups are used for access rights"
	}
	Else
	{
		Line 0 "Active Directory groups are not used for access rights"
	}
	Line 0 ""
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): "
}
#endregion

#region process PVS Site functions
Function ProcessPVSSite
{
	#build site values
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Processing Sites"
	$GetWhat = "site"
	$GetParam = ""
	$ErrorTxt = "PVS Site information"
	$PVSSites = BuildPVSObject $GetWhat $GetParam $ErrorTxt
	
	If($Null -eq $PVSSites)
	{
		Write-Host -foregroundcolor Red -backgroundcolor Black "WARNING: $(Get-Date -Format G): No Sites Found"
		Line 0 "No Sites Found "
	}
	Else
	{
		ForEach($PVSSite in $PVSSites)
		{
			Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): `tProcessing Site $($PVSSite.siteName)"
			Line 0 "Site Name: " $PVSSite.siteName

			#security tab
			Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): `t`tProcessing Security Tab"
			$temp = $PVSSite.SiteName
			$GetWhat = "authgroup"
			$GetParam = "sitename = $temp"
			$ErrorTxt = "Groups with Site Administrator access"
			$authgroups = BuildPVSObject $GetWhat $GetParam $ErrorTxt
			If($Null -ne $authGroups)
			{
				Line 1 "Groups with Site Administrator access`t`t:"
				ForEach($Group in $authgroups)
				{
					Line 7 "  " $Group.authGroupName
				}
			}
			Else
			{
				Line 1 "Groups with Site Administrator access`t`t: No Site Administrators defined"
			}

			#MAK tab
			#MAK User and Password are encrypted

			#options tab
			Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): `t`tProcessing Options Tab"
			If($PVSVersion -eq "5" -or (($PVSVersion -eq "6" -or $PVSVersion -eq "7") -and $FarmAutoAddEnabled))
			{
				Line 1 "Add new devices to this collection`t`t: " -nonewline
				If($PVSSite.DefaultCollectionName)
				{
					Line 0 $PVSSite.DefaultCollectionName
				}
				Else
				{
					Line 0 "<No Default collection>"
				}
			}
			If($PVSVersion -eq "6" -or $PVSVersion -eq "7")
			{
				If($PVSVersion -eq "6")
				{
					Line 1 "Seconds between vDisk inventory scans`t`t: " $PVSSite.inventoryFilePollingInterval
				}

				#vDisk Update
				Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): `t`tProcessing vDisk Update Tab"
				If($PVSSite.enableDiskUpdate -eq "1")
				{
					Line 1 "Enable automatic vDisk updates on this site`t: Yes"
					Line 1 "Server to run vDisk updates for this site`t: " $PVSSite.diskUpdateServerName
				}
				Else
				{
					Line 1 "Enable automatic vDisk updates on this site`t: No"
				}
			}
			Line 0 ""
			
			#process all servers in site
			Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): `t`tProcessing Servers in Site $($PVSSite.siteName)"
			$temp = $PVSSite.SiteName
			$GetWhat = "server"
			$GetParam = "sitename = $temp"
			$ErrorTxt = "Servers for Site $temp"
			$servers = BuildPVSObject $GetWhat $GetParam $ErrorTxt
			
			If($Null -eq $servers)
			{
				Write-Host -foregroundcolor Red -backgroundcolor Black "WARNING: $(Get-Date -Format G): `t`tNo Servers Found in Site $($PVSSite.siteName)"
				Line 0 "No Servers Found in Site $($PVSSite.siteName)"
			}
			Else
			{
				Line 1 "Servers"
				ForEach($Server in $Servers)
				{
					#first make sure the SOAP service is running on the server
					If(VerifyPVSSOAPService $Server.serverName)
					{
						Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): `t`t`tProcessing Server $($Server.serverName)"
						#general tab
						Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): `t`t`t`tProcessing General Tab"
						Line 2 "Name`t`t`t`t`t`t: " $Server.serverName
						Line 2 "Log events to the server's Windows Event Log`t: " -nonewline
						If($Server.eventLoggingEnabled -eq "1")
						{
							Line 0 "Yes"
						}
						Else
						{
							Line 0 "No"
							$obj1 = [PSCustomObject] @{
								ItemText = "$($Server.serverName) event logging is not enabled"
							}
							$null = $Script:ItemsToReview.Add($obj1)
						}
							
						Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): `t`t`t`tProcessing Network Tab"
						$test = $Server.ip.ToString()
						$test1 = $test.replace(",",", ")
						
						$tmparray= @($server.ip.split(","))
						
						ForEach($item in $tmparray)
						{
							$obj1 = [PSCustomObject] @{
								ServerName = $Server.serverName							
								IPAddress  = $item							
							}
							$null = $Script:StreamingIPAddresses.Add($obj1)
						}
						If($Script:PVSVersion -eq "7")
						{
							Line 2 "Streaming IP addresses`t`t`t`t: " $test1
						}
						Else
						{
							Line 2 "IP addresses`t`t`t`t`t: " $test1
						}
						Line 2 "First port`t`t`t`t`t: " $Server.firstPort
						Line 2 "Last port`t`t`t`t`t: " $Server.lastPort
						If($Script:PVSVersion -eq "7")
						{
							Line 2 "Management IP`t`t`t`t`t: " $Server.managementIp
							$obj1 = [PSCustomObject] @{
								ServerName = $Server.serverName							
								IPAddress  = $Server.managementIp
							}
							$Script:NICIPAddresses.Add( $Server.serverName, $Server.managementIp )
						}
							
						#create array for appendix A
						
						Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): `t`t`t`t`tGather Advanced server info"
						$obj1 = [PSCustomObject] @{
							ServerName              = $Server.serverName						
							ThreadsPerPort          = $Server.threadsPerPort						
							BuffersPerThread        = $Server.buffersPerThread						
							ServerCacheTimeout      = $Server.serverCacheTimeout						
							LocalConcurrentIOLimit  = $Server.localConcurrentIoLimit						
							RemoteConcurrentIOLimit = $Server.remoteConcurrentIoLimit						
							maxTransmissionUnits    = $Server.maxTransmissionUnits						
							IOBurstSize             = $Server.ioBurstSize						
							NonBlockingIOEnabled    = $Server.nonBlockingIoEnabled						
						}
						$null = $Script:AdvancedItems1.Add($obj1)
						
						$obj2 = [PSCustomObject] @{
							ServerName              = $Server.serverName						
							BootPauseSeconds        = $Server.bootPauseSeconds						
							MaxBootSeconds          = $Server.maxBootSeconds						
							MaxBootDevicesAllowed   = $Server.maxBootDevicesAllowed						
							vDiskCreatePacing       = $Server.vDiskCreatePacing						
							LicenseTimeout          = $Server.licenseTimeout						
						}
						$null = $Script:AdvancedItems2.Add($obj2)
						
						GetComputerWMIInfo $server.ServerName
							
						GetConfigWizardInfo $server.ServerName
							
						GetDisableTaskOffloadInfo $server.ServerName
							
						GetBootstrapInfo $server
							
						GetPVSServiceInfo $server.ServerName

						GetBadStreamingIPAddresses $server.ServerName
						
						GetMiscRegistryKeys $server.ServerName
						
						GetMicrosoftHotfixes $server.ServerName
						
						GetInstalledRolesAndFeatures $server.ServerName
						
						GetPVSProcessInfo $server.ServerName
						
						GetCitrixInstalledComponents $server.ServerName
					}
					Else
					{
						Line 2 "Name: " $Server.serverName
						Line 2 "Server was not processed because the server was offLine or the SOAP Service was not running"
						Line 0 ""
					}
				}
			}

			#process all device collections in site
			Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): `t`tProcessing all device collections in site"
			$Temp = $PVSSite.SiteName
			$GetWhat = "Collection"
			$GetParam = "siteName = $Temp"
			$ErrorTxt = "Device Collection information"
			$Collections = BuildPVSObject $GetWhat $GetParam $ErrorTxt

			If($Null -ne $Collections)
			{
				Line 1 "Device Collections"
				ForEach($Collection in $Collections)
				{
					Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): `t`t`tProcessing Collection $($Collection.collectionName)"
					Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): `t`t`t`tProcessing General Tab"
					Line 2 "Name: " $Collection.collectionName

					Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): `t`t`t`tProcessing Security Tab"
					$Temp = $Collection.collectionId
					$GetWhat = "authGroup"
					$GetParam = "collectionId = $Temp"
					$ErrorTxt = "Device Collection information"
					$AuthGroups = BuildPVSObject $GetWhat $GetParam $ErrorTxt

					$DeviceAdmins = $False
					If($Null -ne $AuthGroups)
					{
						Line 3 "Groups with 'Device Administrator' access`t:"
						ForEach($AuthGroup in $AuthGroups)
						{
							$Temp = $authgroup.authGroupName
							$GetWhat = "authgroupusage"
							$GetParam = "authgroupname = $Temp"
							$ErrorTxt = "Device Collection Administrator usage information"
							$AuthGroupUsages = BuildPVSObject $GetWhat $GetParam $ErrorTxt
							If($Null -ne $AuthGroupUsages)
							{
								ForEach($AuthGroupUsage in $AuthGroupUsages)
								{
									If($AuthGroupUsage.role -eq "300")
									{
										$DeviceAdmins = $True
										Line 9 "  " $authgroup.authGroupName
									}
								}
							}
						}
					}
					If(!$DeviceAdmins)
					{
						Line 3 "Groups with 'Device Administrator' access`t: None defined"
					}

					$DeviceOperators = $False
					If($Null -ne $AuthGroups)
					{
						Line 3 "Groups with 'Device Operator' access`t`t:"
						ForEach($AuthGroup in $AuthGroups)
						{
							$Temp = $authgroup.authGroupName
							$GetWhat = "authgroupusage"
							$GetParam = "authgroupname = $Temp"
							$ErrorTxt = "Device Collection Operator usage information"
							$AuthGroupUsages = BuildPVSObject $GetWhat $GetParam $ErrorTxt
							If($Null -ne $AuthGroupUsages)
							{
								ForEach($AuthGroupUsage in $AuthGroupUsages)
								{
									If($AuthGroupUsage.role -eq "400")
									{
										$DeviceOperators = $True
										Line 9 "  " $authgroup.authGroupName
									}
								}
							}
						}
					}
					If(!$DeviceOperators)
					{
						Line 3 "Groups with 'Device Operator' access`t`t: None defined"
					}

					Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): `t`t`t`tProcessing Auto-Add Tab"
					If($Script:FarmAutoAddEnabled)
					{
						Line 3 "Template target device`t`t`t`t: " $Collection.templateDeviceName
						If(![String]::IsNullOrEmpty($Collection.autoAddPrefix) -or ![String]::IsNullOrEmpty($Collection.autoAddPrefix))
						{
							Line 4 "Device Name"
						}
						If(![String]::IsNullOrEmpty($Collection.autoAddPrefix))
						{
							Line 5 "Prefix`t`t`t: " $Collection.autoAddPrefix
						}
						Line 5 "Length`t`t`t: " $Collection.autoAddNumberLength
						Line 5 "Zero fill`t`t: " -nonewline
						If($Collection.autoAddZeroFill -eq "1")
						{
							Line 0 "Yes"
						}
						Else
						{
							Line 0 "No"
						}
						If(![String]::IsNullOrEmpty($Collection.autoAddPrefix))
						{
							Line 5 "Suffix`t`t`t: " $Collection.autoAddSuffix
						}
						Line 5 "Last incremental #`t: " $Collection.lastAutoAddDeviceNumber
					}
					Else
					{
						Line 3 "The auto-add feature is not enabled at the PVS Farm level"
					}
					#for each collection process each device
					Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): `t`t`t`tProcessing the first device in each collection"
					$Temp = $Collection.collectionId
					$GetWhat = "deviceInfo"
					$GetParam = "collectionId = $Temp"
					$ErrorTxt = "Device Info information"
					$Devices = BuildPVSObject $GetWhat $GetParam $ErrorTxt
					
					If($Null -ne $Devices)
					{
						Line 0 ""
						$Device = $Devices[0]
						Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): `t`t`t`t`tProcessing Device $($Device.deviceName)"
						If($Device.type -eq "3")
						{
							Line 3 "Device with Personal vDisk Properties"
						}
						Else
						{
							Line 3 "Target Device Properties"
						}
						Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): `t`t`t`t`t`tProcessing General Tab"
						Line 4 "Name`t`t`t`t`t: " $Device.deviceName
						If(($PVSVersion -eq "6" -or $PVSVersion -eq "7") -and $Device.type -ne "3")
						{
							Line 4 "Type`t`t`t`t`t: " -nonewline
							Switch ($Device.type)
							{
								0 {Line 0 "Production"; Break}
								1 {Line 0 "Test"; Break}
								2 {Line 0 "Maintenance"; Break}
								3 {Line 0 "Personal vDisk"; Break}
								Default {Line 0 "Device type could not be determined: $($Device.type)"; Break}
							}
						}
						If($Device.type -ne "3")
						{
							Line 4 "Boot from`t`t`t`t: " -nonewline
							Switch ($Device.bootFrom)
							{
								1 {Line 0 "vDisk"; Break}
								2 {Line 0 "Hard Disk"; Break}
								3 {Line 0 "Floppy Disk"; Break}
								Default {Line 0 "Boot from could not be determined: $($Device.bootFrom)"; Break}
							}
						}
						Line 4 "Port`t`t`t`t`t: " $Device.port
						If($Device.type -ne "3")
						{
							Line 4 "Disabled`t`t`t`t: " -nonewline
							If($Device.enabled -eq "1")
							{
								Line 0 "No"
							}
							Else
							{
								Line 0 "Yes"
							}
						}
						Else
						{
							Line 4 "vDisk`t`t`t`t`t: " $Device.diskLocatorName
							Line 4 "Personal vDisk Drive`t`t`t: " $Device.pvdDriveLetter
						}
						Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): `t`t`t`t`t`tProcessing vDisks Tab"
						#process all vdisks for this device
						$Temp = $Device.deviceName
						$GetWhat = "DiskInfo"
						$GetParam = "deviceName = $Temp"
						$ErrorTxt = "Device vDisk information"
						$vDisks = BuildPVSObject $GetWhat $GetParam $ErrorTxt
						If($Null -ne $vDisks)
						{
							ForEach($vDisk in $vDisks)
							{
								Line 4 "vDisk Name`t`t`t`t: $($vDisk.storeName)`\$($vDisk.diskLocatorName)"
							}
						}
						Line 4 "List local hard drive in boot menu`t: " -nonewline
						If($Device.localDiskEnabled -eq "1")
						{
							Line 0 "Yes"
						}
						Else
						{
							Line 0 "No"
						}
						
						DeviceStatus $Device
					}
					Else
					{
						Line 3 "No Target Devices found. Device Collection is empty."
						Line 0 ""
						$obj1 = [PSCustomObject] @{
							CollectionName = $Collection.collectionName
						}
						$null = $Script:EmptyDeviceCollections.Add($obj1)
					}
				}
			}
		}
	}

	#Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): "
}

Function VerifyPVSSOAPService
{
	Param([string]$PVSServer='')
	
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): `t`t`tVerifying server $($PVSServer) is online"
	If(Test-Connection -ComputerName $PVSServer -quiet -EA 0)
	{

		Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): `t`t`tVerifying PVS SOAP Service is running on server $($PVSServer)"
		$soapserver = $Null

		$soapserver = Get-Service -ComputerName $PVSServer -EA 0 | Where-Object {$_.Name -like "soapserver"}

		If($soapserver.Status -ne "Running")
		{
			Write-Warning "The Citrix PVS Soap Server service is not Started on server $($PVSServer)"
			Write-Warning "Server $($PVSServer) cannot be processed.  See message above."
			Return $False
		}
		Else
		{
			Return $True
		}
	}
	Else
	{
		Write-Warning "The server $($PVSServer) is offLine or unreachable."
		Write-Warning "Server $($PVSServer) cannot be processed.  See message above."
		Return $False
	}
}

Function GetInstalledRolesAndFeatures
{
	Param([string]$ComputerName)
	
	#don't do for server 2008 r2 because get-windowsfeature doesn't support -computername
	If($Script:RunningOS -like "*2008*")
	{
		#don't do anything
	}
	Else
	{
		#added V1.16 get Windows installed Roles and Features
		Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): `t`t`t`tRetrieving Windows installed Roles and Features"
		[bool]$GotWinComponents = $True
		
		$results = Get-WindowsFeature -ComputerName $ComputerName -EA 0 4> $Null
		
		If(!$?)
		{
			$GotWinComponents = $False
		}
		
		$WinComponents = $results | Where-Object Installed | Select-Object DisplayName,Name,FeatureType | Sort-Object DisplayName 
		
		If($GotWinComponents -eq $False)
		{
			Line 1 "No Windows installed Roles and Features were found"
			Line 0 ""
		}
		Else
		{
			ForEach($Component in $WinComponents)
			{
				$obj1 = [PSCustomObject] @{
					DisplayName	= $Component.DisplayName			
					Name		= $Component.Name			
					ServerName	= $ComputerName			
					FeatureType	= $Component.FeatureType			
				}
				$null = $Script:WinInstalledComponents.Add($obj1)
			}
		}
	}
}

Function DeviceStatus
{
	Param($xDevice)

	If($Null -eq $xDevice -or $xDevice.status -eq "" -or $xDevice.status -eq "0")
	{
		Line 3 "Target device inactive"
	}
	Else
	{
		Line 2 "Target device active"
		Line 3 "IP Address: " $xDevice.ip
		Line 3 "Server: $($xDevice.serverName)"
		Line 3 "Server IP: $($xDevice.serverIpConnection)"
		Line 3 "Server Port: $($xDevice.serverPortConnection)"
		Line 3 "vDisk: " $xDevice.diskLocatorName
		Line 3 "vDisk version: " $xDevice.diskVersion
		Line 3 "vDisk name: " $xDevice.diskFileName
		Line 3 "vDisk access: " -nonewline
		Switch ($xDevice.diskVersionAccess)
		{
			0 {Line 0 "Production"; Break}
			1 {Line 0 "Test"; Break}
			2 {Line 0 "Maintenance"; Break}
			3 {Line 0 "Personal vDisk"; Break}
			Default {Line 0 "vDisk access type could not be determined: $($xDevice.diskVersionAccess)"; Break}
		}
		If($PVSVersion -eq "7")
		{
			Line 3 "Local write cache disk:$($xDevice.localWriteCacheDiskSize)GB"
			Line 3 "Boot mode:" -nonewline
			Switch($xDevice.bdmBoot)
			{
				0 {Line 0 "PXE boot"; Break}
				1 {Line 0 "BDM disk"; Break}
				Default {Line 0 "Boot mode could not be determined: $($xDevice.bdmBoot)"; Break}
			}
		}
		Switch($xDevice.licenseType)
		{
			0 {Line 3 "No License"; Break}
			1 {Line 3 "Desktop License"; Break}
			2 {Line 3 "Server License"; Break}
			5 {Line 3 "OEM SmartClient License"; Break}
			6 {Line 3 "XenApp License"; Break}
			7 {Line 3 "XenDesktop License"; Break}
			Default {Line 0 "Device license type could not be determined: $($xDevice.licenseType)"; Break}
		}
		
		Line 3 "Logging level: " -nonewline
		Switch ($xDevice.logLevel)
		{
			0   {Line 0 "Off"; Break}
			1   {Line 0 "Fatal"; Break}
			2   {Line 0 "Error"; Break}
			3   {Line 0 "Warning"; Break}
			4   {Line 0 "Info"; Break}
			5   {Line 0 "Debug"; Break}
			6   {Line 0 "Trace"; Break}
			Default {Line 0 "Logging level could not be determined: $($xDevice.logLevel)"; Break}
		}
	}
	Line 0 ""
}

Function GetBootstrapInfo
{
	Param([object]$server)

	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): `t`t`t`tProcessing Bootstrap files"
	Line 2 "Bootstrap settings"
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): `t`t`t`t`tProcessing Bootstrap files for Server $($server.servername)"
	#first get all bootstrap files for the server
	$temp = $server.serverName
	$GetWhat = "ServerBootstrapNames"
	$GetParam = "serverName = $temp"
	$ErrorTxt = "Server Bootstrap Name information"
	$BootstrapNames = BuildPVSObject $GetWhat $GetParam $ErrorTxt

	#Now that the list of bootstrap names has been gathered
	#We have the mandatory parameter to get the bootstrap info
	#there should be at least one bootstrap filename
	If($Null -ne $Bootstrapnames)
	{
		#cannot use the BuildPVSObject Function here
		$serverbootstraps = @()
		ForEach($Bootstrapname in $Bootstrapnames)
		{
			#get serverbootstrap info
			$error.Clear()
			$tempserverbootstrap = Mcli-Get ServerBootstrap -p name="$($Bootstrapname.name)",servername="$($server.serverName)"
			If($error.Count -eq 0)
			{
				$serverbootstrap = $Null
				ForEach($record in $tempserverbootstrap)
				{
					If($record.length -gt 5 -and $record.substring(0,6) -eq "Record")
					{
						If($Null -ne $serverbootstrap)
						{
							$serverbootstraps +=  $serverbootstrap
						}
						$serverbootstrap = new-object System.Object
						#add the bootstrapname name value to the serverbootstrap object
						$property = "BootstrapName"
						$value = $Bootstrapname.name
						Add-Member -inputObject $serverbootstrap -MemberType NoteProperty -Name $property -Value $value
					}
					$index = $record.IndexOf(':')
					If($index -gt 0)
					{
						$property = $record.SubString(0, $index)
						$value = $record.SubString($index + 2)
						If($property -ne "Executing")
						{
							Add-Member -inputObject $serverbootstrap -MemberType NoteProperty -Name $property -Value $value
						}
					}
				}
				$serverbootstraps +=  $serverbootstrap
			}
			Else
			{
				Line 2 "Server Bootstrap information could not be retrieved"
				Line 2 "Error returned is " $error[0].FullyQualifiedErrorId.Split(',')[0].Trim()
			}
		}
		If($Null -ne $ServerBootstraps)
		{
			ForEach($ServerBootstrap in $ServerBootstraps)
			{
				Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): `t`t`t`t`tProcessing Bootstrap file $($ServerBootstrap.Bootstrapname)"
				$obj1 = [PSCustomObject] @{
					ServerName 	  = $Server.serverName				
					BootstrapName = $ServerBootstrap.Bootstrapname				
					IP1        	  = $ServerBootstrap.bootserver1_Ip				
					IP2        	  = $ServerBootstrap.bootserver2_Ip				
					IP3        	  = $ServerBootstrap.bootserver3_Ip				
					IP4        	  = $ServerBootstrap.bootserver4_Ip				
				}
				$null = $Script:BootstrapItems.Add($obj1)

				Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): `t`t`t`t`t`tProcessing Bootstrap General Tab"
				Line 3 "Bootstrap file`t: " $ServerBootstrap.Bootstrapname
				If($ServerBootstrap.bootserver1_Ip -ne "0.0.0.0")
				{
					Line 3 "IP Address`t: " $ServerBootstrap.bootserver1_Ip
					Line 3 "Subnet Mask`t: " $ServerBootstrap.bootserver1_Netmask
					Line 3 "Gateway`t`t: " $ServerBootstrap.bootserver1_Gateway
					Line 3 "Port`t`t: " $ServerBootstrap.bootserver1_Port
				}
				If($ServerBootstrap.bootserver2_Ip -ne "0.0.0.0")
				{
					Line 3 "IP Address`t: " $ServerBootstrap.bootserver2_Ip
					Line 3 "Subnet Mask`t: " $ServerBootstrap.bootserver2_Netmask
					Line 3 "Gateway`t`t: " $ServerBootstrap.bootserver2_Gateway
					Line 3 "Port`t`t: " $ServerBootstrap.bootserver2_Port
				}
				If($ServerBootstrap.bootserver3_Ip -ne "0.0.0.0")
				{
					Line 3 "IP Address`t: " $ServerBootstrap.bootserver3_Ip
					Line 3 "Subnet Mask`t: " $ServerBootstrap.bootserver3_Netmask
					Line 3 "Gateway`t`t: " $ServerBootstrap.bootserver3_Gateway
					Line 3 "Port`t`t: " $ServerBootstrap.bootserver3_Port
				}
				If($ServerBootstrap.bootserver4_Ip -ne "0.0.0.0")
				{
					Line 3 "IP Address`t: " $ServerBootstrap.bootserver4_Ip
					Line 3 "Subnet Mask`t: " $ServerBootstrap.bootserver4_Netmask
					Line 3 "Gateway`t`t: " $ServerBootstrap.bootserver4_Gateway
					Line 3 "Port`t`t: " $ServerBootstrap.bootserver4_Port
				}
				Line 0 ""
				Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): `t`t`t`t`t`tProcessing Bootstrap Options Tab"
				Line 3 "Verbose mode`t`t: " -nonewline
				If($ServerBootstrap.verboseMode -eq "1")
				{
					Line 0 "Yes"
				}
				Else
				{
					Line 0 "No"
				}
				Line 3 "Interrupt safe mode`t: " -nonewline
				If($ServerBootstrap.interruptSafeMode -eq "1")
				{
					Line 0 "Yes"
				}
				Else
				{
					Line 0 "No"
				}
				Line 3 "Advanced Memory Support`t: " -nonewline
				If($ServerBootstrap.paeMode -eq "1")
				{
					Line 0 "Yes"
				}
				Else
				{
					Line 0 "No"
				}
				Line 3 "Network recovery method`t: " -nonewline
				If($ServerBootstrap.bootFromHdOnFail -eq "0")
				{
					Line 0 "Restore network connection"
				}
				Else
				{
					Line 0 "Reboot to Hard Drive after $($ServerBootstrap.recoveryTime) seconds"
				}
				Line 3 "Login polling timeout`t: " -nonewline
				If($ServerBootstrap.pollingTimeout -eq "")
				{
					Line 0 "5000 (milliseconds)"
				}
				Else
				{
					Line 0 "$($ServerBootstrap.pollingTimeout) (milliseconds)"
				}
				Line 3 "Login general timeout`t: " -nonewline
				If($ServerBootstrap.generalTimeout -eq "")
				{
					Line 0 "5000 (milliseconds)"
				}
				Else
				{
					Line 0 "$($ServerBootstrap.generalTimeout) (milliseconds)"
				}
				Line 0 ""
			}
		}
	}
	Else
	{
		Line 2 "No Bootstrap names available"
	}
	Line 0 ""
}

Function GetPVSServiceInfo
{
	Param([string]$ComputerName)

	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): `t`t`t`tProcessing PVS Services for Server $($server.servername)"
	$Services = Get-WmiObject -ComputerName $ComputerName Win32_Service -EA 0 | `
	Where-Object {$_.DisplayName -like "Citrix PVS*"} | `
	Select-Object displayname, name, status, startmode, started, startname, state | `
	Sort-Object DisplayName
	
	If($? -and $Null -ne $Services)
	{
		ForEach($Service in $Services)
		{
			$obj1 = [PSCustomObject] @{
				ServerName 	   = $ComputerName
				DisplayName	   = $Service.DisplayName
				Name  		   = $Service.Name
				Status 		   = $Service.Status
				StartMode  	   = $Service.StartMode
				Started		   = $Service.Started
				StartName  	   = $Service.StartName
				State  		   = $Service.State
				FailureAction1 = "Take no Action"
				FailureAction2 = "Take no Action"
				FailureAction3 = "Take no Action"
			}

			[array]$Actions = sc.exe \\$ComputerName qfailure $Service.Name
			
			If($Actions.Length -gt 0)
			{
				If(($Actions -like "*RESTART -- Delay*") -or ($Actions -like "*RUN PROCESS -- Delay*") -or ($Actions -like "*REBOOT -- Delay*"))
				{
					$cnt = 0
					ForEach($Item in $Actions)
					{
						Switch ($Item)
						{
							{$Item -like "*RESTART -- Delay*"}		{$cnt++; $obj1.$("FailureAction$($Cnt)") = "Restart the Service"; Break}
							{$Item -like "*RUN PROCESS -- Delay*"}	{$cnt++; $obj1.$("FailureAction$($Cnt)") = "Run a Program"; Break}
							{$Item -like "*REBOOT -- Delay*"}		{$cnt++; $obj1.$("FailureAction$($Cnt)") = "Restart the Computer"; Break}
						}
					}
				}
			}
			
			$null = $Script:PVSServiceItems.Add($obj1)
		}
	}
	Else
	{
		Line 2 "No PVS services found for $($ComputerName)"
	}
	Line 0 ""
}

Function GetPVSProcessInfo
{
	Param([string]$ComputerName)
	
	#Whether or not the Inventory executable is running (Inventory.exe)
	#Whether or not the Notifier executable is running (Notifier.exe)
	#Whether or not the MgmtDaemon executable is running (MgmtDaemon.exe)
	#Whether or not the StreamProcess executable is running (StreamProcess.exe)
	
	#All four of those run within the StreamService.exe process.

	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): `t`t`t`tRetrieving PVS Processes for Server $($server.servername)"

	Try
	{
		$InventoryProcess = Get-Process -Name 'Inventory' -ComputerName $ComputerName

		$tmp1 = "Inventory"
		$tmp2 = ""
		If($InventoryProcess)
		{
			$tmp2 = "Running"
		}
		Else
		{
			$tmp2 = "Not Running"
		}
		$obj1 = [PSCustomObject] @{
			ProcessName	= $tmp1
			ServerName 	= $ComputerName	
			Status  	= $tmp2
		}
		$null = $Script:PVSProcessItems.Add($obj1)
	}
	
	Catch
	{
		$tmp1 = "Inventory"
		$tmp2 = "Unable to retrieve"
		$obj1 = [PSCustomObject] @{
			ProcessName	= $tmp1
			ServerName 	= $ComputerName	
			Status  	= $tmp2
		}
		$null = $Script:PVSProcessItems.Add($obj1)
	}
	
	Try
	{
		$NotifierProcess = Get-Process -Name 'Notifier' -ComputerName $ComputerName

		$tmp1 = "Notifier"
		$tmp2 = ""
		If($NotifierProcess)
		{
			$tmp2 = "Running"
		}
		Else
		{
			$tmp2 = "Not Running"
		}
		$obj1 = [PSCustomObject] @{
			ProcessName	= $tmp1
			ServerName 	= $ComputerName	
			Status  	= $tmp2
		}
		$null = $Script:PVSProcessItems.Add($obj1)
	}
	
	Catch
	{
		$tmp1 = "Notifier"
		$tmp2 = "Unable to retrieve"
		$obj1 = [PSCustomObject] @{
			ProcessName	= $tmp1
			ServerName 	= $ComputerName	
			Status  	= $tmp2
		}
		$null = $Script:PVSProcessItems.Add($obj1)
	}
	
	Try
	{
		$MgmtDaemonProcess = Get-Process -Name 'MgmtDaemon' -ComputerName $ComputerName
	
		$tmp1 = "MgmtDaemon"
		$tmp2 = ""
		If($MgmtDaemonProcess)
		{
			$tmp2 = "Running"
		}
		Else
		{
			$tmp2 = "Not Running"
		}
		$obj1 = [PSCustomObject] @{
			ProcessName	= $tmp1
			ServerName 	= $ComputerName	
			Status  	= $tmp2
		}
		$null = $Script:PVSProcessItems.Add($obj1)
	}
	
	Catch
	{
		$tmp1 = "MgmtDaemon"
		$tmp2 = "Unable to retrieve"
		$obj1 = [PSCustomObject] @{
			ProcessName	= $tmp1
			ServerName 	= $ComputerName	
			Status  	= $tmp2
		}
		$null = $Script:PVSProcessItems.Add($obj1)
	}
	
	Try
	{
		$StreamProcessProcess = Get-Process -Name 'StreamProcess' -ComputerName $ComputerName
	
		$tmp1 = "StreamProcess"
		$tmp2 = ""
		If($StreamProcessProcess)
		{
			$tmp2 = "Running"
		}
		Else
		{
			$tmp2 = "Not Running"
		}
		$obj1 = [PSCustomObject] @{
			ProcessName	= $tmp1
			ServerName 	= $ComputerName	
			Status  	= $tmp2
		}
		$null = $Script:PVSProcessItems.Add($obj1)
	}
	
	Catch
	{
		$tmp1 = "StreamProcess"
		$tmp2 = "Unable to retrieve"
		$obj1 = [PSCustomObject] @{
			ProcessName	= $tmp1
			ServerName 	= $ComputerName	
			Status  	= $tmp2
		}
		$null = $Script:PVSProcessItems.Add($obj1)
	}
}

Function GetBadStreamingIPAddresses
{
	Param([string]$ComputerName)
	#function updated by Andrew Williamson @ Fujitsu Services to handle servers with multiple NICs
	#further optimization by Michael B. Smith

	#loop through the configured streaming ip address and compare to the physical configured ip addresses
	#if a streaming ip address is not in the list of physical ip addresses, it is a bad streaming ip address
	ForEach ($Stream in ($Script:StreamingIPAddresses | Where-Object {$_.Servername -eq $ComputerName})) {
		$exists = $false
		:outerLoop ForEach ($ServerNIC in $Script:NICIPAddresses.Item($ComputerName)) 
		{
			ForEach ($IP in $ServerNIC) 
			{ 
				# there could be more than one IP
				If ($Stream.IPAddress -eq $IP) 
				{
					$Exists = $true
					break :outerLoop
				}
			}
		}
		if (!$exists) 
		{
			$obj1 = [PSCustomObject] @{
				ServerName = $ComputerName			
				IPAddress  = $Stream.IPAddress			
			}
			$null = $Script:BadIPs.Add($obj1)
		}
	}
}

Function GetMicrosoftHotfixes 
{
	Param([string]$ComputerName)
	
	#added V1.16 get installed Microsoft Hotfixes and Updates
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): `t`t`t`tRetrieving Microsoft hotfixes and updates"
	[bool]$GotMSHotfixes = $True
	
	Try
	{
		$results = Get-HotFix -computername $ComputerName | Select-Object CSName,Caption,Description,HotFixID,InstalledBy,InstalledOn
		$MSInstalledHotfixes = $results | Sort-Object HotFixID
		$results = $Null
	}
	
	Catch
	{
		$GotMSHotfixes = $False
	}

	If($GotMSHotfixes -eq $False)
	{
		Line 1 "No installed Microsoft hotfixes or updates were found"
		Line 0 ""
	}
	Else
	{
		ForEach($Hotfix in $MSInstalledHotfixes)
		{
			$obj1 = [PSCustomObject] @{
				HotFixID	= $Hotfix.HotFixID			
				ServerName	= $Hotfix.CSName			
				Caption		= $Hotfix.Caption			
				Description	= $Hotfix.Description			
				InstalledBy	= $Hotfix.InstalledBy			
				InstalledOn	= $Hotfix.InstalledOn			
			}
			$null = $Script:MSHotfixes.Add($obj1)
		}
	}
}

Function GetCitrixInstalledComponents 
{
	Param([string]$ComputerName)
	
	#added V1.24 get installed Citrix components
	#code adapted from the CVAD doc script
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): `t`t`t`tRetrieving Citrix installed components"
	[bool]$GotCtxComponents = $True
	
	If($ComputerName -eq $env:computername)
	{
		$results = Get-ChildItem HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall|`
		ForEach-Object{Get-ItemProperty $_.pspath}|`
		Where-Object { $_.PSObject.Properties[ 'Publisher' ] -and $_.Publisher -like 'Citrix*'}|`
		Select-Object DisplayName, DisplayVersion
	}
	Else
	{
		#see if the remote registy service is running
		$serviceresults = Get-Service -ComputerName $ComputerName -Name "RemoteRegistry" -EA 0
		If($? -and $Null -ne $serviceresults)
		{
			If($serviceresults.Status -eq "Running")
			{
				$results = Invoke-Command -ComputerName $ComputerName -ScriptBlock `
				{Get-ChildItem HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall|`
				ForEach-Object{Get-ItemProperty $_.pspath}|`
				Where-Object { $_.PSObject.Properties[ 'Publisher' ] -and $_.Publisher -like 'Citrix*'}|`
				Select-Object DisplayName, DisplayVersion}
			}
		}
		Else
		{
			$results = $Null
			$GotCtxComponents = $False
		}
	}
	
	If(!$? -or $Null -eq $results)
	{
		$GotCtxComponents = $False
	}
	Else
	{
		$CtxComponents = $results
		$results = $Null
		
		If($ComputerName -eq $env:computername)
		{
			$results = Get-ChildItem HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall|`
			ForEach-Object{Get-ItemProperty $_.pspath}|`
			Where-Object { $_.PSObject.Properties[ 'Publisher' ] -and $_.Publisher -like 'Citrix*'}|`
			Select-Object DisplayName, DisplayVersion
		}
		Else
		{
			$results = Invoke-Command -ComputerName $ComputerName -ScriptBlock `
			{Get-ChildItem HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall|`
			ForEach-Object{Get-ItemProperty $_.pspath}|`
			Where-Object { $_.PSObject.Properties[ 'Publisher' ] -and $_.Publisher -like 'Citrix*'}|`
			Select-Object DisplayName, DisplayVersion}
		}
		If($?)
		{
			$CtxComponents += $results
		}
		
		$CtxComponents = $CtxComponents | Sort-Object DisplayName
	}
	
	If($GotCtxComponents)
	{
		ForEach($Component in $CtxComponents)
		{
			$obj1 = [PSCustomObject] @{
				DisplayName    = $Component.DisplayName						
				DisplayVersion = $Component.DisplayVersion						
				PVSServerName  = $ComputerName						
			}
			$null = $Script:CtxInstalledComponents.Add($obj1)
		}
	}
	Else
	{
		Line 1 "No Citrix Installed Components were found"
	}
}

Function GetComputerWMIInfo
{
	Param([string]$RemoteComputerName)
	
	# original work by Kees Baggerman, 
	# Senior Technical Consultant @ Inter Access
	# k.baggerman@myvirtualvision.com
	# @kbaggerman on Twitter
	# http://blog.myvirtualvision.com
	# modified 1-May-2014 to work in trusted AD Forests and using different domain admin credentials	
	# modified 17-Aug-2016 to fix a few issues with Text and HTML output
	# modified 2-Apr-2018 to add ComputerOS information

	#Get Computer info
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): `t`t`t`tProcessing WMI Computer information"
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): `t`t`t`t`tHardware information"
	Line 0 ""
	Line 2 "Computer Information: $($RemoteComputerName)"
	Line 3 "General Computer"
	
	Try
	{
		$Results = Get-WmiObject -computername $RemoteComputerName win32_computersystem
	}
	
	Catch
	{
		$Results = $Null
	}
	
	If($? -and $Null -ne $Results)
	{
		$ComputerItems = $Results | Select-Object Manufacturer, Model, Domain, `
		@{N="TotalPhysicalRam"; E={[math]::round(($_.TotalPhysicalMemory / 1GB),0)}}, `
		NumberOfProcessors, NumberOfLogicalProcessors
		$Results = $Null
		[string]$ComputerOS = (Get-WmiObject -class Win32_OperatingSystem -computername $RemoteComputerName -EA 0).Caption

		ForEach($Item in $ComputerItems)
		{
			OutputComputerItem $Item $ComputerOS $RemoteComputerName
		}
	}
	ElseIf(!$?)
	{
		Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Get-WmiObject win32_computersystem failed for $($RemoteComputerName)"
		Write-Warning "Get-WmiObject win32_computersystem failed for $($RemoteComputerName)"
		Line 4 "Get-WmiObject win32_computersystem failed for $($RemoteComputerName)"
		Line 4 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository"
		Line 4 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may"
		Line 4 "need to rerun the script with Domain Admin credentials from the trusted Forest."
		Line 4 ""
	}
	Else
	{
		Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): No results Returned for Computer information"
		Line 4 "No results Returned for Computer information"
	}
	
	#Get Disk info
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): `t`t`t`t`tDrive information"

	Line 3 "Drive(s)"

	Try
	{
		$Results = Get-WmiObject -computername $RemoteComputerName Win32_LogicalDisk
	}
	
	Catch
	{
		$Results = $Null
	}

	If($? -and $Null -ne $Results)
	{
		$drives = $Results | Select-Object caption, @{N="drivesize"; E={[math]::round(($_.size / 1GB),0)}}, 
		filesystem, @{N="drivefreespace"; E={[math]::round(($_.freespace / 1GB),0)}}, 
		volumename, drivetype, volumedirty, volumeserialnumber
		$Results = $Null
		ForEach($drive in $drives)
		{
			If($drive.caption -ne "A:" -and $drive.caption -ne "B:")
			{
				OutputDriveItem $drive $RemoteComputerName
			}
		}
	}
	ElseIf(!$?)
	{
		Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Get-WmiObject Win32_LogicalDisk failed for $($RemoteComputerName)"
		Write-Warning "Get-WmiObject Win32_LogicalDisk failed for $($RemoteComputerName)"
		Line 4 "Get-WmiObject Win32_LogicalDisk failed for $($RemoteComputerName)"
		Line 4 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository"
		Line 4 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may"
		Line 4 "need to rerun the script with Domain Admin credentials from the trusted Forest."
	}
	Else
	{
		Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): No results Returned for Drive information"
		Line 4 "No results Returned for Drive information"
	}
	

	#Get CPU's and stepping
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): `t`t`t`t`tProcessor information"

	Line 3 "Processor(s)"

	Try
	{
		$Results = Get-WmiObject -computername $RemoteComputerName win32_Processor
	}
	
	Catch
	{
		$Results = $Null
	}

	If($? -and $Null -ne $Results)
	{
		$Processors = $Results | Select-Object availability, name, description, maxclockspeed, 
		l2cachesize, l3cachesize, numberofcores, numberoflogicalprocessors
		$Results = $Null
		ForEach($processor in $processors)
		{
			OutputProcessorItem $processor $RemoteComputerName
		}
	}
	ElseIf(!$?)
	{
		Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Get-WmiObject win32_Processor failed for $($RemoteComputerName)"
		Write-Warning "Get-WmiObject win32_Processor failed for $($RemoteComputerName)"
		Line 4 "Get-WmiObject win32_Processor failed for $($RemoteComputerName)"
		Line 4 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository"
		Line 4 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may"
		Line 4 "need to rerun the script with Domain Admin credentials from the trusted Forest."
	}
	Else
	{
		Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): No results Returned for Processor information"
		Line 4 "No results Returned for Processor information"
	}

	#Get Nics
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): `t`t`t`t`tNIC information"

	Line 3 "Network Interface(s)"

	[bool]$GotNics = $True
	
	Try
	{
		$Results = Get-WmiObject -computername $RemoteComputerName win32_networkadapterconfiguration
	}
	
	Catch
	{
		$Results = $Null
	}

	If($? -and $Null -ne $Results)
	{
		$Nics = $Results | Where-Object {$Null -ne $_.ipaddress}
		$Results = $Null

		If($Null -eq $Nics) 
		{ 
			$GotNics = $False 
		} 
		Else 
		{ 
			$GotNics = !($Nics.__PROPERTY_COUNT -eq 0) 
		} 
	
		If($GotNics)
		{
			ForEach($nic in $nics)
			{
				Try
				{
					$ThisNic = Get-WmiObject -computername $RemoteComputerName win32_networkadapter | Where-Object {$_.index -eq $nic.index}
				}
				
				Catch 
				{
					$ThisNic = $Null
				}
				
				If($? -and $Null -ne $ThisNic)
				{
					OutputNicItem $Nic $ThisNic $RemoteComputerName
				}
				ElseIf(!$?)
				{
					Write-Warning "$(Get-Date -Format G): Error retrieving NIC information"
					Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)"
					Write-Warning "Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)"
					Line 4 "Error retrieving NIC information"
					Line 4 "Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)"
					Line 4 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository"
					Line 4 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may"
					Line 4 "need to rerun the script with Domain Admin credentials from the trusted Forest."
				}
				Else
				{
					Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): No results Returned for NIC information"
					Line 4 "No results Returned for NIC information"
				}
			}
		}	
	}
	ElseIf(!$?)
	{
		Write-Warning "$(Get-Date -Format G): Error retrieving NIC configuration information"
		Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)"
		Write-Warning "Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)"
		Line 4 "Error retrieving NIC configuration information"
		Line 4 "Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)"
		Line 4 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository"
		Line 4 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may"
		Line 4 "need to rerun the script with Domain Admin credentials from the trusted Forest."
	}
	Else
	{
		Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): No results Returned for NIC configuration information"
		Line 4 "No results Returned for NIC configuration information"
	}
	
	Line 0 ""
}

Function OutputComputerItem
{
	Param([object]$Item, [string]$OS, [string]$RemoteComputerName)
	# modified 2-Apr-2018 to add Operating System information
	
	#get computer's power plan
	#https://techcommunity.microsoft.com/t5/core-infrastructure-and-security/get-the-active-power-plan-of-multiple-servers-with-powershell/ba-p/370429
	
	try 
	{

		$PowerPlan = (Get-WmiObject -ComputerName $RemoteComputerName -Class Win32_PowerPlan -Namespace "root\cimv2\power" |
			Where-Object {$_.IsActive -eq $true} |
			Select-Object @{Name = "PowerPlan"; Expression = {$_.ElementName}}).PowerPlan
	}

	catch 
	{

		$PowerPlan = $_.Exception

	}	
	
	Line 4 "Manufacturer`t`t`t: " $Item.manufacturer
	Line 4 "Model`t`t`t`t: " $Item.model
	Line 4 "Domain`t`t`t`t: " $Item.domain
	Line 4 "Operating System`t`t: " $OS
	Line 4 "Power Plan`t`t`t: " $PowerPlan
	Line 4 "Total Ram`t`t`t: $($Item.totalphysicalram) GB"
	Line 4 "Physical Processors (sockets)`t: " $Item.NumberOfProcessors
	Line 4 "Logical Processors (cores w/HT)`t: " $Item.NumberOfLogicalProcessors
	Line 4 ""
	$obj1 = [PSCustomObject] @{
		ServerName         = $RemoteComputerName
		OperatingSystem    = $OS
		PowerPlan          = $PowerPlan
		TotalRam           = $Item.totalphysicalram
		PhysicalProcessors = $Item.NumberOfProcessors
		LogicalProcessors  = $Item.NumberOfLogicalProcessors
	}
	$null = $Script:ServerComputerItemsToReview.Add($obj1)
}

Function OutputDriveItem
{
	Param([object]$Drive, [string]$RemoteComputerName)
	
	$xDriveType = ""
	Switch ($drive.drivetype)
	{
		0	{$xDriveType = "Unknown"; Break}
		1	{$xDriveType = "No Root Directory"; Break}
		2	{$xDriveType = "Removable Disk"; Break}
		3	{$xDriveType = "Local Disk"; Break}
		4	{$xDriveType = "Network Drive"; Break}
		5	{$xDriveType = "Compact Disc"; Break}
		6	{$xDriveType = "RAM Disk"; Break}
		Default {$xDriveType = "Unknown"; Break}
	}
	
	$xVolumeDirty = ""
	If(![String]::IsNullOrEmpty($drive.volumedirty))
	{
		If($drive.volumedirty)
		{
			$xVolumeDirty = "Yes"
		}
		Else
		{
			$xVolumeDirty = "No"
		}
	}

	Line 4 "Caption`t`t: " $drive.caption
	Line 4 "Size`t`t: $($drive.drivesize) GB"
	If(![String]::IsNullOrEmpty($drive.filesystem))
	{
		Line 4 "File System`t: " $drive.filesystem
	}
	Line 4 "Free Space`t: $($drive.drivefreespace) GB"
	If(![String]::IsNullOrEmpty($drive.volumename))
	{
		Line 4 "Volume Name`t: " $drive.volumename
	}
	If(![String]::IsNullOrEmpty($drive.volumedirty))
	{
		Line 4 "Volume is Dirty`t: " $xVolumeDirty
	}
	If(![String]::IsNullOrEmpty($drive.volumeserialnumber))
	{
		Line 4 "Volume Serial #`t: " $drive.volumeserialnumber
	}
	Line 4 "Drive Type`t: " $xDriveType
	Line 4 ""
	$obj1 = [PSCustomObject] @{
		ServerName   = $RemoteComputerName
		DriveCaption = $drive.caption
		DriveSize    = $drive.drivesize
	}
	$null = $Script:ServerDriveItemsToReview.Add($obj1)
}

Function OutputProcessorItem
{
	Param([object]$Processor)
	
	$xAvailability = ""
	Switch ($processor.availability)
	{
		1	{$xAvailability = "Other"; Break }
		2	{$xAvailability = "Unknown"; Break }
		3	{$xAvailability = "Running or Full Power"; Break }
		4	{$xAvailability = "Warning"; Break }
		5	{$xAvailability = "In Test"; Break }
		6	{$xAvailability = "Not Applicable"; Break }
		7	{$xAvailability = "Power Off"; Break }
		8	{$xAvailability = "Off Line"; Break }
		9	{$xAvailability = "Off Duty"; Break }
		10	{$xAvailability = "Degraded"; Break }
		11	{$xAvailability = "Not Installed"; Break }
		12	{$xAvailability = "Install Error"; Break }
		13	{$xAvailability = "Power Save - Unknown"; Break }
		14	{$xAvailability = "Power Save - Low Power Mode"; Break }
		15	{$xAvailability = "Power Save - Standby"; Break }
		16	{$xAvailability = "Power Cycle"; Break }
		17	{$xAvailability = "Power Save - Warning"; Break }
		Default	{$xAvailability = "Unknown"; Break }
	}

	Line 4 "Name`t`t`t`t: " $processor.name
	Line 4 "Description`t`t`t: " $processor.description
	Line 4 "Max Clock Speed`t`t`t: $($processor.maxclockspeed) MHz"
	If($processor.l2cachesize -gt 0)
	{
		Line 4 "L2 Cache Size`t`t`t: $($processor.l2cachesize) KB"
	}
	If($processor.l3cachesize -gt 0)
	{
		Line 4 "L3 Cache Size`t`t`t: $($processor.l3cachesize) KB"
	}
	If($processor.numberofcores -gt 0)
	{
		Line 4 "# of Cores`t`t`t: " $processor.numberofcores
	}
	If($processor.numberoflogicalprocessors -gt 0)
	{
		Line 4 "# of Logical Procs (cores w/HT)`t: " $processor.numberoflogicalprocessors
	}
	Line 4 "Availability`t`t`t: " $xAvailability
	Line 4 ""
	$obj1 = [PSCustomObject] @{
		ServerName   = $RemoteComputerName
		Cores        = $processor.numberofcores
		LogicalProcs = $processor.numberoflogicalprocessors
	}
	$null = $Script:ServerProcessorItemsToReview.Add($obj1)
}

Function OutputNicItem
{
	Param([object]$Nic, [object]$ThisNic, [string]$RemoteComputerName)
	
	$powerMgmt = Get-WmiObject -computername $RemoteComputerName MSPower_DeviceEnable -Namespace root\wmi | Where-Object{$_.InstanceName -match [regex]::Escape($ThisNic.PNPDeviceID)}

	If($? -and $Null -ne $powerMgmt)
	{
		If($powerMgmt.Enable -eq $True)
		{
			$PowerSaving = "Enabled"
		}
		Else
		{
			$PowerSaving = "Disabled"
		}
	}
	Else
	{
        $PowerSaving = "N/A"
	}
	
	$xAvailability = ""
	Switch ($ThisNic.availability)
	{
		1		{$xAvailability = "Other"; Break}
		2		{$xAvailability = "Unknown"; Break}
		3		{$xAvailability = "Running or Full Power"; Break}
		4		{$xAvailability = "Warning"; Break}
		5		{$xAvailability = "In Test"; Break}
		6		{$xAvailability = "Not Applicable"; Break}
		7		{$xAvailability = "Power Off"; Break}
		8		{$xAvailability = "Off Line"; Break}
		9		{$xAvailability = "Off Duty"; Break}
		10		{$xAvailability = "Degraded"; Break}
		11		{$xAvailability = "Not Installed"; Break}
		12		{$xAvailability = "Install Error"; Break}
		13		{$xAvailability = "Power Save - Unknown"; Break}
		14		{$xAvailability = "Power Save - Low Power Mode"; Break}
		15		{$xAvailability = "Power Save - Standby"; Break}
		16		{$xAvailability = "Power Cycle"; Break}
		17		{$xAvailability = "Power Save - Warning"; Break}
		Default	{$xAvailability = "Unknown"; Break}
	}

	#attempt to get Receive Side Scaling setting
	$RSSEnabled = "N/A"
	Try
	{
		#https://ios.developreference.com/article/10085450/How+do+I+enable+VRSS+(Virtual+Receive+Side+Scaling)+for+a+Windows+VM+without+relying+on+Enable-NetAdapterRSS%3F
		$RSSEnabled = (Get-WmiObject -ComputerName $RemoteComputerName MSFT_NetAdapterRssSettingData -Namespace "root\StandardCimV2" -ea 0).Enabled

		If($RSSEnabled)
		{
			$rssenabled = "Enabled"
		}
		ELse
		{
			$rssenabled = "Disabled"
		}
	}
	
	Catch
	{
		$RSSEnabled = "Not available on $Script:RunningOS"
	}
	
	$xIPAddress = @()
	ForEach($IPAddress in $Nic.ipaddress)
	{
		$xIPAddress += "$($IPAddress)"
	}

	$xIPSubnet = @()
	ForEach($IPSubnet in $Nic.ipsubnet)
	{
		$xIPSubnet += "$($IPSubnet)"
	}

	If($Null -ne $nic.dnsdomainsuffixsearchorder -and $nic.dnsdomainsuffixsearchorder.length -gt 0)
	{
		$nicdnsdomainsuffixsearchorder = $nic.dnsdomainsuffixsearchorder
		$xnicdnsdomainsuffixsearchorder = @()
		ForEach($DNSDomain in $nicdnsdomainsuffixsearchorder)
		{
			$xnicdnsdomainsuffixsearchorder += "$($DNSDomain)"
		}
	}
	
	If($Null -ne $nic.dnsserversearchorder -and $nic.dnsserversearchorder.length -gt 0)
	{
		$nicdnsserversearchorder = $nic.dnsserversearchorder
		$xnicdnsserversearchorder = @()
		ForEach($DNSServer in $nicdnsserversearchorder)
		{
			$xnicdnsserversearchorder += "$($DNSServer)"
		}
	}

	$xdnsenabledforwinsresolution = ""
	If($nic.dnsenabledforwinsresolution)
	{
		$xdnsenabledforwinsresolution = "Yes"
	}
	Else
	{
		$xdnsenabledforwinsresolution = "No"
	}
	
	$xTcpipNetbiosOptions = ""
	Switch ($nic.TcpipNetbiosOptions)
	{
		0		{$xTcpipNetbiosOptions = "Use NetBIOS setting from DHCP Server"; Break}
		1		{$xTcpipNetbiosOptions = "Enable NetBIOS"; Break}
		2		{$xTcpipNetbiosOptions = "Disable NetBIOS"; Break}
		Default	{$xTcpipNetbiosOptions = "Unknown"; Break}
	}
	
	$xwinsenablelmhostslookup = ""
	If($nic.winsenablelmhostslookup)
	{
		$xwinsenablelmhostslookup = "Yes"
	}
	Else
	{
		$xwinsenablelmhostslookup = "No"
	}

	Line 4 "Name`t`t`t`t: " $ThisNic.Name
	If($ThisNic.Name -ne $nic.description)
	{
		Line 4 "Description`t`t`t: " $nic.description
	}
	Line 4 "Connection ID`t`t`t: " $ThisNic.NetConnectionID
	Line 4 "Manufacturer`t`t`t: " $ThisNic.manufacturer
	Line 4 "Availability`t`t`t: " $xAvailability
    Line 4 "Allow the computer to turn off "
	Line 4 "this device to save power`t: " $PowerSaving
	Line 4 "Receive Side Scaling`t`t: " $RSSEnabled
	Line 4 "Physical Address`t`t: " $nic.macaddress
	Line 4 "IP Address`t`t`t: " $xIPAddress[0]
	$cnt = -1
	ForEach($tmp in $xIPAddress)
	{
		$cnt++
		If($cnt -gt 0)
		{
			Line 8 "  " $tmp
		}
	}
	Line 4 "Default Gateway`t`t`t: " $Nic.Defaultipgateway
	Line 4 "Subnet Mask`t`t`t: " $xIPSubnet[0]
	$cnt = -1
	ForEach($tmp in $xIPSubnet)
	{
		$cnt++
		If($cnt -gt 0)
		{
			Line 8 "  " $tmp
		}
	}
	If($nic.dhcpenabled)
	{
		$DHCPLeaseObtainedDate = $nic.ConvertToDateTime($nic.dhcpleaseobtained)
		$DHCPLeaseExpiresDate = $nic.ConvertToDateTime($nic.dhcpleaseexpires)
		Line 4 "DHCP Enabled`t`t`t: " $nic.dhcpenabled
		Line 4 "DHCP Lease Obtained`t`t: " $dhcpleaseobtaineddate
		Line 4 "DHCP Lease Expires`t`t: " $dhcpleaseexpiresdate
		Line 4 "DHCP Server`t`t`t:" $nic.dhcpserver
	}
	If(![String]::IsNullOrEmpty($nic.dnsdomain))
	{
		Line 4 "DNS Domain`t`t`t: " $nic.dnsdomain
	}
	If($Null -ne $nic.dnsdomainsuffixsearchorder -and $nic.dnsdomainsuffixsearchorder.length -gt 0)
	{
		[int]$x = 1
		Line 4 "DNS Search Suffixes`t`t: " $xnicdnsdomainsuffixsearchorder[0]
		$cnt = -1
		ForEach($tmp in $xnicdnsdomainsuffixsearchorder)
		{
			$cnt++
			If($cnt -gt 0)
			{
				Line 8 "  " $tmp
			}
		}
	}
	Line 4 "DNS WINS Enabled`t`t: " $xdnsenabledforwinsresolution
	If($Null -ne $nic.dnsserversearchorder -and $nic.dnsserversearchorder.length -gt 0)
	{
		[int]$x = 1
		Line 4 "DNS Servers`t`t`t: " $xnicdnsserversearchorder[0]
		$cnt = -1
		ForEach($tmp in $xnicdnsserversearchorder)
		{
			$cnt++
			If($cnt -gt 0)
			{
				Line 8 "  " $tmp
			}
		}
	}
	Line 4 "NetBIOS Setting`t`t`t: " $xTcpipNetbiosOptions
	Line 4 "Enabled LMHosts`t`t`t: " $xwinsenablelmhostslookup
	If(![String]::IsNullOrEmpty($nic.winshostlookupfile))
	{
		Line 4 "Host Lookup File`t`t: " $nic.winshostlookupfile
	}
	If(![String]::IsNullOrEmpty($nic.winsprimaryserver))
	{
		Line 4 "Primary Server`t`t`t: " $nic.winsprimaryserver
	}
	If(![String]::IsNullOrEmpty($nic.winssecondaryserver))
	{
		Line 4 "Secondary Server`t`t: " $nic.winssecondaryserver
	}
	If(![String]::IsNullOrEmpty($nic.winsscopeid))
	{
		Line 4 "Scope ID`t`t`t: " $nic.winsscopeid
	}
	Line 0 ""
	$obj1 = [PSCustomObject] @{
		ServerName   = $RemoteComputerName
		Name         = $ThisNic.Name
		Manufacturer = $ThisNic.manufacturer
		PowerMgmt    = $PowerSaving
		RSS          = $RSSEnabled
	}
	$null = $Script:ServerNICItemsToReview.Add($obj1)
}
#endregion

#region Process vDisks in Farm functions
Function ProcessvDisksinFarm
{
	#process all vDisks in site
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): `t`tProcessing all vDisks in site"
	[int]$NumberofvDisks = 0
	$GetWhat = "DiskInfo"
	$GetParam = ""
	$ErrorTxt = "Disk information"
	$Disks = BuildPVSObject $GetWhat $GetParam $ErrorTxt

	Line 0 "vDisks in Farm"
	If($Null -ne $Disks)
	{
		ForEach($Disk in $Disks)
		{
			Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): `t`t`tProcessing vDisk $($Disk.diskLocatorName)"
			Line 1 $Disk.diskLocatorName
			If($Script:PVSVersion -eq "5")
			{
				#PVS 5.x
				Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): `t`t`t`tProcessing General Tab"
				Line 2 "Store: " $Disk.storeName
				Line 2 "Site: " $Disk.siteName
				Line 2 "Filename: " $Disk.diskLocatorName
				Line 2 "Size: " (($Disk.diskSize/1024)/1024)/1024 -nonewline
				Line 0 " MB"
				If(![String]::IsNullOrEmpty($Disk.serverName))
				{
					Line 2 "Use this server to provide the vDisk: " $Disk.serverName
				}
				Else
				{
					Line 2 "Subnet Affinity: " -nonewline
					Switch ($Disk.subnetAffinity)
					{
						0 {Line 0 "None"; Break}
						1 {Line 0 "Best Effort"; Break}
						2 {Line 0 "Fixed"; Break}
						Default {Line 2 "Subnet Affinity could not be determined: $($Disk.subnetAffinity)"; Break}
					}
					Line 2 "Rebalance Enabled: " -nonewline
					If($Disk.rebalanceEnabled -eq "1")
					{
						Line 0 "Yes"
						Line 2 "Trigger Percent: $($Disk.rebalanceTriggerPercent)"
					}
					Else
					{
						Line 0 "No"
					}
				}
				Line 2 "Allow use of this vDisk: " -nonewline
				If($Disk.enabled -eq "1")
				{
					Line 0 "Yes"
					If($Disk.deviceCount -gt 0)
					{
						$NumberofvDisks++
					}
				}
				Else
				{
					Line 0 "No"
				}
				Line 2 "Access mode: " -nonewline
				If($Disk.writeCacheType -eq "0")
				{
					Line 0 "Private Image (single device, read/write access)"
				}
				ElseIf($Disk.writeCacheType -eq "7")
				{
					Line 0 "Difference Disk Image"
				}
				Else
				{
					Line 0 "Standard Image (multi-device, read-only access)"
					Line 2 "Cache type: " -nonewline
					Switch ($Disk.writeCacheType)
					{
						0   {Line 0 "Private Image"; Break}
						1   {
								Line 0 "Cache on server"
								
								$obj1 = [PSCustomObject] @{
									StoreName = $Disk.storeName								
									SiteName  = $Disk.siteName								
									vDiskName = $Disk.diskLocatorName								
								}
								$null = $Script:CacheOnServer.Add($obj1)
								Break
							}
						2   {Line 0 "Cache encrypted on server disk"; Break}
						3   {
							Line 0 "Cache in device RAM"
							Line 2 "Cache Size: $($Disk.writeCacheSize) MBs"; Break
							}
						4   {Line 0 "Cache on device's HD"; Break}
						5   {Line 0 "Cache encrypted on device's hard disk"; Break}
						6   {Line 0 "RAM Disk"; Break}
						7   {Line 0 "Difference Disk"; Break}
						Default {Line 0 "Cache type could not be determined: $($Disk.writeCacheType)"; Break}
					}
				}
				If($Disk.activationDateEnabled -eq "0")
				{
					Line 2 "Enable automatic updates for the vDisk: " -nonewline
					If($Disk.autoUpdateEnabled -eq "1")
					{
						Line 0 "Yes"
					}
					Else
					{
						Line 0 "No"
					}
					Line 2 "Apply vDisk updates as soon as they are detected by the server"
				}
				Else
				{
					Line 2 "Enable automatic updates for the vDisk: " -nonewline
					If($Disk.autoUpdateEnabled -eq "1")
					{
						Line 0 "Yes"
					}
					Else
					{
						Line 0 "No"
					}
					Line 2 "Schedule the next vDisk update to occur on: $($Disk.activeDate)"
				}
				Line 2 "Microsoft license type: " -nonewline
				Switch ($Disk.licenseMode)
				{
					0 {Line 0 "None"; Break}
					1 {Line 0 "Multiple Activation Key (MAK)"; Break}
					2 {Line 0 "Key Management Service (KMS)"; Break}
					Default {Line 0 "Volume License Mode could not be determined: $($Disk.licenseMode)"; Break}
				}
				#options tab
				Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): `t`t`t`tProcessing Options Tab"
				Line 2 "High availability (HA): " -nonewline
				If($Disk.haEnabled -eq "1")
				{
					Line 0 "Yes"
				}
				Else
				{
					Line 0 "No"
				}
				Line 2 "AD machine account password management: " -nonewline
				If($Disk.adPasswordEnabled -eq "1")
				{
					Line 0 "Yes"
				}
				Else
				{
					Line 0 "No"
				}
				
				Line 2 "Printer management: " -nonewline
				If($Disk.printerManagementEnabled -eq "1")
				{
					Line 0 "Yes"
				}
				Else
				{
					Line 0 "No"
				}
			}
			Else
			{
				#PVS 6.x or 7.x
				Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): `t`t`tProcessing vDisk Properties"
				Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): `t`t`t`tProcessing General Tab"
				Line 2 "Site`t`t`t`t`t`t: " $Disk.siteName
				Line 2 "Store`t`t`t`t`t`t: " $Disk.storeName
				Line 2 "Filename`t`t`t`t`t: " $Disk.diskLocatorName
				Line 2 "Size`t`t`t`t`t`t: " (($Disk.diskSize/1024)/1024)/1024 -nonewline
				Line 0 " MB"
				Line 2 "VHD block size`t`t`t`t`t: " $Disk.vhdBlockSize -nonewline
				Line 0 " KB"
				Line 2 "Access mode`t`t`t`t`t: " -nonewline
				If($Disk.writeCacheType -eq "0")
				{
					Line 0 "Private Image (single device, read/write access)"
				}
				Else
				{
					Line 0 "Standard Image (multi-device, read-only access)"
					Line 2 "Cache type: " -nonewline
					Switch ($Disk.writeCacheType)
					{
						0   {Line 0 "Private Image"; Break}
						1   {
								Line 0 "Cache on server"
								
								$obj1 = [PSCustomObject] @{
									StoreName = $Disk.storeName								
									SiteName  = $Disk.siteName								
									vDiskName = $Disk.diskLocatorName								
								}
								$null = $Script:CacheOnServer.Add($obj1)
								Break
							}
						3   {
							Line 0 "Cache in device RAM"
							Line 2 "Cache Size: $($Disk.writeCacheSize) MBs"; Break
							}
						4   {Line 0 "Cache on device's hard disk"; Break}
						6   {Line 0 "RAM Disk"; Break}
						7   {Line 0 "Difference Disk"; Break}
						8  	{Line 0 "Cache on device hard drive persisted (NT 6.1 and later)"; Break}
						9  	{Line 0 "Cache in device RAM with overflow on hard disk"; Break}
						10 	{Line 0 "Private Image with Asynchronous IO"; Break} #added 1811
						11 	{Line 0 "Cache on server, persistent with Asynchronous IO"; Break} #added 1811
						12 	{Line 0 "Cache in device RAM with overflow on hard disk with Asynchronous IO"; Break} #added 1811
						Default {Line 0 "Cache type could not be determined: $($Disk.writeCacheType)"; Break}
					}
				}
				If(![String]::IsNullOrEmpty($Disk.menuText))
				{
					Line 2 "BIOS boot menu text`t`t`t`t: " $Disk.menuText
				}
				Line 2 "Enable AD machine account password management`t: " -nonewline
				If($Disk.adPasswordEnabled -eq "1")
				{
					Line 0 "Yes"
				}
				Else
				{
					Line 0 "No"
				}
				
				Line 2 "Enable printer management`t`t`t: " -nonewline
				If($Disk.printerManagementEnabled -eq "1")
				{
					Line 0 "Yes"
				}
				Else
				{
					Line 0 "No"
				}
				Line 2 "Enable streaming of this vDisk`t`t`t: " -nonewline
				If($Disk.Enabled -eq "1")
				{
					Line 0 "Yes"
					If($Disk.deviceCount -gt 0)
					{
						$NumberofvDisks++
					}
				}
				Else
				{
					Line 0 "No"
				}
				Line 2 "Microsoft license type`t`t`t`t: " -nonewline
				Switch ($Disk.licenseMode)
				{
					0 {Line 0 "None"; Break}
					1 {Line 0 "Multiple Activation Key (MAK)"; Break}
					2 {Line 0 "Key Management Service (KMS)"; Break}
					Default {Line 0 "Volume License Mode could not be determined: $($Disk.licenseMode)"; Break}
				}
				If($Script:PVSFullVersion -ge "7.22")
				{
					If($Disk.AccelerateOfficeActivation)
					{
						Line 2 "Accelerate Office Activation`t`t`t: Yes"
					}
					Else
					{
						Line 2 "Accelerate Office Activation`t`t`t: No"
					}
				}

				Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): `t`t`t`tProcessing Auto Update Tab"
				If($Disk.activationDateEnabled -eq "0")
				{
					Line 2 "Enable automatic updates for the vDisk`t`t: " -nonewline
					If($Disk.autoUpdateEnabled -eq "1")
					{
						Line 0 "Yes"
					}
					Else
					{
						Line 0 "No"
					}
					Line 2 "Apply vDisk updates as soon as they are detected by the server"
				}
				Else
				{
					Line 2 "Enable automatic updates for the vDisk`t`t: " -nonewline
					If($Disk.autoUpdateEnabled -eq "1")
					{
						Line 0 "Yes"
					}
					Else
					{
						Line 0 "No"
					}
					Line 2 "Schedule the next vDisk update to occur on`t: $($Disk.activeDate)"
				}
				#process Versions menu
				#get versions info
				#thanks to the PVS Product team for their help in understanding the Versions information
				Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): `t`t`tProcessing vDisk Versions"
				$error.Clear()
				$MCLIGetResult = Mcli-Get DiskVersion -p diskLocatorName="$($Disk.diskLocatorName)",storeName="$($disk.storeName)",siteName="$($disk.siteName)"
				If($error.Count -eq 0)
				{
					#build versions object
					$PluralObject = @()
					$SingleObject = $Null
					ForEach($record in $MCLIGetResult)
					{
						If($record.length -gt 5 -and $record.substring(0,6) -eq "Record")
						{
							If($Null -ne $SingleObject)
							{
								$PluralObject += $SingleObject
							}
							$SingleObject = new-object System.Object
						}

						$index = $record.IndexOf(':')
						If($index -gt 0)
						{
							$property = $record.SubString(0, $index)
							$value    = $record.SubString($index + 2)
							If($property -ne "Executing")
							{
								Add-Member -inputObject $SingleObject -MemberType NoteProperty -Name $property -Value $value
							}
						}
					}
					$PluralObject += $SingleObject
					$DiskVersions = $PluralObject
					
					If($Null -ne $DiskVersions)
					{
						#get the current booting version
						#by default, the $DiskVersions object is in version number order lowest to highest
						#the initial or base version is 0 and always exists
						[string]$BootingVersion = "0"
						[bool]$BootOverride = $False
						ForEach($DiskVersion in $DiskVersions)
						{
							If($DiskVersion.access -eq "3")
							{
								#override i.e. manually selected boot version
								$BootingVersion = $DiskVersion.version
								$BootOverride = $True
								Break
							}
							ElseIf($DiskVersion.access -eq "0" -and $DiskVersion.IsPending -eq "0" )
							{
								$BootingVersion = $DiskVersion.version
								$BootOverride = $False
							}
						}
						
						Line 2 "Boot production devices from version`t`t: " -NoNewLine
						If($BootOverride)
						{
							Line 0 $BootingVersion
						}
						Else
						{
							Line 0 "Newest released"
						}
						Line 0 ""
						
						$VersionFlag = $False
						ForEach($DiskVersion in $DiskVersions)
						{
							Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): `t`t`t`tProcessing vDisk Version $($DiskVersion.version)"
							Line 2 "Version`t`t`t`t`t`t: " -NoNewLine
							If($DiskVersion.version -eq $BootingVersion)
							{
								Line 0 "$($DiskVersion.version) (Current booting version)"
							}
							Else
							{
								Line 0 $DiskVersion.version
							}
							If($DiskVersion.version -gt $Script:farm.maxVersions -and $VersionFlag -eq $False)
							{
								$VersionFlag = $True
								Line 2 "Version of vDisk is $($DiskVersion.version) which is greater than the limit of $($Script:farm.maxVersions). Consider merging."
								
								$obj1 = [PSCustomObject] @{
									vDiskName = $Disk.diskLocatorName								
								}
								$null = $Script:VersionsToMerge.Add($obj1)
							}
							Line 2 "Created`t`t`t`t`t`t: " $DiskVersion.createDate
							If(![String]::IsNullOrEmpty($DiskVersion.scheduledDate))
							{
								Line 2 "Released`t`t`t`t`t: " $DiskVersion.scheduledDate
							}
							Line 2 "Devices`t`t`t`t`t`t: " $DiskVersion.deviceCount
							Line 2 "Access`t`t`t`t`t`t: " -NoNewLine
							Switch ($DiskVersion.access)
							{
								"0" {Line 0 "Production"; Break}
								"1" {Line 0 "Maintenance"; Break}
								"2" {Line 0 "Maintenance Highest Version"; Break}
								"3" {Line 0 "Override"; Break}
								"4" {Line 0 "Merge"; Break}
								"5" {Line 0 "Merge Maintenance"; Break}
								"6" {Line 0 "Merge Test"; Break}
								"7" {Line 0 "Test"; Break}
								Default {Line 0 "Access could not be determined: $($DiskVersion.access)"; Break}
							}
							Line 2 "Type`t`t`t`t`t`t: " -NoNewLine
							Switch ($DiskVersion.type)
							{
								"0" {Line 0 "Base"; Break}
								"1" {Line 0 "Manual"; Break}
								"2" {Line 0 "Automatic"; Break}
								"3" {Line 0 "Merge"; Break}
								"4" {Line 0 "Merge Base"; Break}
								Default {Line 0 "Type could not be determined: $($DiskVersion.type)"; Break}
							}
							If(![String]::IsNullOrEmpty($DiskVersion.description))
							{
								Line 2 "Properties`t`t`t`t`t: " $DiskVersion.description
							}
							Line 2 "Can Delete`t`t`t`t`t: "  -NoNewLine
							Switch ($DiskVersion.canDelete)
							{
								0 {Line 0 "No"; Break}
								1 {Line 0 "Yes"; Break}
							}
							Line 2 "Can Merge`t`t`t`t`t: "  -NoNewLine
							Switch ($DiskVersion.canMerge)
							{
								0 {Line 0 "No"; Break}
								1 {Line 0 "Yes"; Break}
							}
							Line 2 "Can Merge Base`t`t`t`t`t: "  -NoNewLine
							Switch ($DiskVersion.canMergeBase)
							{
								0 {Line 0 "No"; Break}
								1 {Line 0 "Yes"; Break}
							}
							Line 2 "Can Promote`t`t`t`t`t: "  -NoNewLine
							Switch ($DiskVersion.canPromote)
							{
								0 {Line 0 "No"; Break}
								1 {Line 0 "Yes"; Break}
							}
							Line 2 "Can Revert back to Test`t`t`t`t: "  -NoNewLine
							Switch ($DiskVersion.canRevertTest)
							{
								0 {Line 0 "No"; Break}
								1 {Line 0 "Yes"; Break}
							}
							Line 2 "Can Revert back to Maintenance`t`t`t: "  -NoNewLine
							Switch ($DiskVersion.canRevertMaintenance)
							{
								0 {Line 0 "No"; Break}
								1 {Line 0 "Yes"; Break}
							}
							Line 2 "Can Set Scheduled Date`t`t`t`t: "  -NoNewLine
							Switch ($DiskVersion.canSetScheduledDate)
							{
								0 {Line 0 "No"; Break}
								1 {Line 0 "Yes"; Break}
							}
							Line 2 "Can Override`t`t`t`t`t: "  -NoNewLine
							Switch ($DiskVersion.canOverride)
							{
								0 {Line 0 "No"; Break}
								1 {Line 0 "Yes"; Break}
							}
							Line 2 "Is Pending`t`t`t`t`t: "  -NoNewLine
							Switch ($DiskVersion.isPending)
							{
								0 {Line 0 "No, version Scheduled Date has occurred"; Break}
								1 {Line 0 "Yes, version Scheduled Date has not occurred"; Break}
							}
							Line 2 "Replication Status`t`t`t`t: " -NoNewLine
							Switch ($DiskVersion.goodInventoryStatus)
							{
								0 {Line 0 "Not available on all servers"; Break}
								1 {Line 0 "Available on all servers"; Break}
								Default {Line 0 "Replication status could not be determined: $($DiskVersion.goodInventoryStatus)"; Break}
							}
							Line 2 "Disk Filename`t`t`t`t`t: " $DiskVersion.diskFileName
							Line 0 ""
						}
					}
				}
				Else
				{
					Line 0 "Disk Version information could not be retrieved"
					Line 0 "Error returned is " $error[0].FullyQualifiedErrorId.Split(',')[0].Trim()
				}
				
				#process vDisk Load Balancing Menu
				Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): `t`t`tProcessing vDisk Load Balancing Menu"
				If(![String]::IsNullOrEmpty($Disk.serverName))
				{
					Line 2 "Use this server to provide the vDisk: " $Disk.serverName
				}
				Else
				{
					Line 2 "Subnet Affinity`t`t`t`t`t: " -nonewline
					Switch ($Disk.subnetAffinity)
					{
						0 {Line 0 "None"; Break}
						1 {Line 0 "Best Effort"; Break}
						2 {Line 0 "Fixed"; Break}
						Default {Line 0 "Subnet Affinity could not be determined: $($Disk.subnetAffinity)"; Break}
					}
					Line 2 "Rebalance Enabled`t`t`t`t: " -nonewline
					If($Disk.rebalanceEnabled -eq "1")
					{
						Line 0 "Yes"
						Line 2 "Trigger Percent: $($Disk.rebalanceTriggerPercent)"
					}
					Else
					{
						Line 0 "No"
					}
				}
			}
			Line 0 ""
		}
	}

	Line 1 "Number of vDisks that are Enabled and have active connections: " $NumberofvDisks
	Line 0 ""
	# http://blogs.citrix.com/2013/07/03/pvs-internals-2-how-to-properly-size-your-memory/
	[decimal]$RecRAM = ((2 + ($NumberofvDisks * 2)) * 1.15)
	$RecRAM = "{0:N0}" -f $RecRAM
	Line 1 "Recommended RAM for each PVS Server using XenDesktop vDisks  : $($RecRAM)GB"
	[decimal]$RecRAM = ((2 + ($NumberofvDisks * 4)) * 1.15)
	$RecRAM = "{0:N0}" -f $RecRAM
	Line 1 "Recommended RAM for each PVS Server using XenApp vDisks      : $($RecRAM)GB"
	[decimal]$RecRAM = ((2 + (($NumberofvDisks * 4) + ($NumberofvDisks * 2))) * 1.15)
	$RecRAM = "{0:N0}" -f $RecRAM
	Line 1 "Recommended RAM for each PVS Server using XA & XD vDisks     : $($RecRAM)GB"
	Line 0 ""
	Line 1 "This script is not able to tell if a vDisk is running XenDesktop or XenApp."
	Line 1 "The RAM calculation is done based on both scenarios. The original formula is:"
	Line 1 "2GB + (#XA_vDisk * 4GB) + (#XD_vDisk * 2GB) + 15% (Buffer)"
	Line 1 'PVS Internals 2 - How to properly size your memory by Martin Zugec'
	Line 1 'https://www.citrix.com/blogs/2013/07/03/pvs-internals-2-how-to-properly-size-your-memory/'
	Line 0 ""
}
#endregion

#region process stores functions
Function ProcessStores
{
	#process the stores now
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): `tProcessing Stores"
	Line 0 "Stores Properties"
	$GetWhat = "Store"
	$GetParam = ""
	$ErrorTxt = "Farm Store information"
	$Stores = BuildPVSObject $GetWhat $GetParam $ErrorTxt
	If($Null -ne $Stores)
	{
		ForEach($Store in $Stores)
		{
			Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): `t`tProcessing Store $($Store.StoreName)"
			Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): `t`t`tProcessing General Tab"
			Line 1 "Name: " $Store.StoreName
			
			Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): `t`t`tProcessing Servers Tab"
			Line 1 "Servers"
			#find the servers (and the site) that serve this store
			$GetWhat = "Server"
			$GetParam = ""
			$ErrorTxt = "Server information"
			$Servers = BuildPVSObject $GetWhat $GetParam $ErrorTxt
			$StoreServers = @()
			If($Null -ne $Servers)
			{
				ForEach($Server in $Servers)
				{
					Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): `t`t`t`tProcessing Server $($Server.serverName)"
					$Temp = $Server.serverName
					$GetWhat = "ServerStore"
					$GetParam = "serverName = $Temp"
					$ErrorTxt = "Server Store information"
					$ServerStore = BuildPVSObject $GetWhat $GetParam $ErrorTxt
					$Providers = $ServerStore | Where-Object {$_.StoreName -eq $Store.Storename}
					If($Providers)
					{
						ForEach ($Provider in $Providers)
						{
							$StoreServers += $Provider.ServerName
						}
					}
				}	
			}
			Line 2 "Servers that provide this store:"
			ForEach($StoreServer in $StoreServers)
			{
				Line 3 $StoreServer
			}

			Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): `t`t`tProcessing Paths Tab"
			Line 1 "Paths"

			#Run through the servers again and test each one for the path
			ForEach ($StoreServer in $StoreServers)
			{
				#next few lines from Guy Leech
                [hashtable]$invokeCommandParameters = @{}
                If( $StoreServer -ne $env:COMPUTERNAME -and $StoreServer -ne "$env:COMPUTERNAME.$env:UserDnsDomain" )
                {
                    $invokeCommandParameters.Add( 'ComputerName' , $StoreServer )
                }
				If(Invoke-Command @invokeCommandParameters `
				    -ScriptBlock { Param( [string]$path ) ; `
				    Test-Path -Path $path -PathType Container -ErrorAction SilentlyContinue } `
				    -ArgumentList $store.path)
				{
					Line 2 "Default store path: $($Store.path) on server $StoreServer is valid"
				}
				Else
				{
					Line 2 "Default store path: $($Store.path) on server $StoreServer is not valid"
				}
			}

			If(![String]::IsNullOrEmpty($Store.cachePath))
			{
				Line 2 "Default write-cache paths: "
				$WCPaths = @($Store.cachePath.Split(","))
				ForEach($StoreServer in $StoreServers)
				{
					ForEach($WCPath in $WCPaths)
					{
						#next few lines from Guy Leech
						[hashtable]$invokeCommandParameters = @{}
						If( $StoreServer -ne $env:COMPUTERNAME -and $StoreServer -ne "$env:COMPUTERNAME.$env:UserDnsDomain" )
						{
							$invokeCommandParameters.Add( 'ComputerName' , $StoreServer )
						}
						If(Invoke-Command @invokeCommandParameters `
							-ScriptBlock { Param( [string]$path ) ; `
							Test-Path -Path $path -PathType Container -ErrorAction SilentlyContinue } `
							-ArgumentList $WCPath)
						{
							Line 3 "Write Cache Path $($WCPath) on server $StoreServer is valid" 
						}
						Else
						{
							Line 3 "Write Cache Path $($WCPath) on server $StoreServer is not valid" 
						}
					}
				}
			}
			Else
			{
				Line 2 "Using the default write-cache path of $($Store.Path)\WriteCache"
			}
			Line 0 ""
		}
	}
	Else
	{
		Line 1 "There are no Stores configured"
	}
	Line 0 ""
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): "
}
#endregion

#region appendix A
Function OutputAppendixA
{
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Create Appendix A Advanced Server Items (Server/Network)"
	#sort the array by servername
	$Script:AdvancedItems1 = $Script:AdvancedItems1 | Sort-Object ServerName
	
	If($CSV)
	{
		$File = "$($Script:pwdpath)\$($Script:farm.FarmName)_HealthCheck_AppendixA_AdvancedServerItems1.csv"
		$Script:AdvancedItems1 | Export-CSV -Force -Encoding ASCII -NoTypeInformation -Path $File
	}

	Line 0 "Appendix A - Advanced Server Items (Server/Network)"
	Line 0 ""
	Line 1 "Server Name      Threads  Buffers  Server   Local       Remote      Ethernet  IO     Enable      "
	Line 1 "                 per      per      Cache    Concurrent  Concurrent  MTU       Burst  Non-blocking"
	Line 1 "                 Port     Thread   Timeout  IO Limit    IO Limit              Size   IO          "
	Line 1 "================================================================================================="

	ForEach($Item in $Script:AdvancedItems1)
	{
		Line 1 ( "{0,-16} {1,-8} {2,-8} {3,-8} {4,-11} {5,-11} {6,-9} {7,-6} {8,-8}" -f `
		$Item.serverName, $Item.threadsPerPort, $Item.buffersPerThread, $Item.serverCacheTimeout, `
		$Item.localConcurrentIoLimit, $Item.remoteConcurrentIoLimit, $Item.maxTransmissionUnits, $Item.ioBurstSize, `
		$Item.nonBlockingIoEnabled )
	}
	Line 0 ""

	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Finished Creating Appendix A - Advanced Server Items (Server/Network)"
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): "
}
#endregion

#region appendix B
Function OutputAppendixB
{
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Create Appendix B Advanced Server Items (Pacing/Device)"
	#sort the array by servername
	$Script:AdvancedItems2 = $Script:AdvancedItems2 | Sort-Object ServerName
	
	If($CSV)
	{
		$File = "$($Script:pwdpath)\$($Script:farm.FarmName)_HealthCheck_AppendixB_AdvancedServerItems2.csv"
		$Script:AdvancedItems2 | Export-CSV -Force -Encoding ASCII -NoTypeInformation -Path $File
	}

	Line 0 "Appendix B - Advanced Server Items (Pacing/Device)"
	Line 0 ""
	Line 1 "Server Name      Boot     Maximum  Maximum  vDisk     License"
	Line 1 "                 Pause    Boot     Devices  Creation  Timeout"
	Line 1 "                 Seconds  Time     Booting  Pacing           "
	Line 1 "============================================================="
	###### "123451234512345  9999999  9999999  9999999  99999999  9999999

	ForEach($Item in $Script:AdvancedItems2)
	{
		Line 1 ( "{0,-16} {1,-8} {2,-8} {3,-8} {4,-9} {5,-8}" -f `
		$Item.serverName, $Item.bootPauseSeconds, $Item.maxBootSeconds, $Item.maxBootDevicesAllowed, `
		$Item.vDiskCreatePacing, $Item.licenseTimeout )
	}
	Line 0 ""

	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Finished Creating Appendix B - Advanced Server Items (Pacing/Device)"
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): "
}
#endregion

#region appendix C
Function OutputAppendixC
{
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Create Appendix C Config Wizard Items"

	#sort the array by servername
	$Script:ConfigWizItems = $Script:ConfigWizItems | Sort-Object ServerName
	
	If($CSV)
	{
		$File = "$($Script:pwdpath)\$($Script:farm.FarmName)_HealthCheck_AppendixC_ConfigWizardItems.csv"
		$Script:ConfigWizItems | Export-CSV -Force -Encoding ASCII -NoTypeInformation -Path $File
	}
	
	Line 0 "Appendix C - Configuration Wizard Settings"
	Line 0 ""
	Line 1 "Server Name      DHCP        PXE        TFTP    User                                               " 
	Line 1 "                 Services    Services   Option  Account                                            "
	Line 1 "================================================================================================"

	If($Script:ConfigWizItems)
	{
		ForEach($Item in $Script:ConfigWizItems)
		{
			Line 1 ( "{0,-16} {1,-11} {2,-9} {3,-7} {4,-50}" -f `
			$Item.serverName, $Item.DHCPServicesValue, $Item.PXEServicesValue, $Item.TFTPOptionValue, `
			$Item.UserAccount )
		}
	}
	Else
	{
		Line 1 "<None found>"
	}
	Line 0 ""
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Finished Creating Appendix C - Config Wizard Items"
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): "
}
#endregion

#region appendix D
Function OutputAppendixD
{
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Create Appendix D Server Bootstrap Items"

	#sort the array by bootstrapname and servername
	$Script:BootstrapItems = $Script:BootstrapItems | Sort-Object BootstrapName, ServerName
	
	If($CSV)
	{
		$File = "$($Script:pwdpath)\$($Script:farm.FarmName)_HealthCheck_AppendixD_ServerBootstrapItems.csv"
		$Script:BootstrapItems | Export-CSV -Force -Encoding ASCII -NoTypeInformation -Path $File
	}
	
	Line 0 "Appendix D - Server Bootstrap Items"
	Line 0 ""
	Line 1 "Bootstrap Name   Server Name      IP1              IP2              IP3              IP4" 
	Line 1 "===================================================================================================="
    ########123456789012345  XXXXXXXXXXXXXXXX 123.123.123.123  123.123.123.123  123.123.123.123  123.123.123.123
	If($Script:BootstrapItems)
	{
		ForEach($Item in $Script:BootstrapItems)
		{
			Line 1 ( "{0,-16} {1,-16} {2,-16} {3,-16} {4,-16} {5,-16}" -f `
			$Item.BootstrapName, $Item.serverName, $Item.IP1, $Item.IP2, $Item.IP3, $Item.IP4 )
		}
	}
	Else
	{
		Line 1 "<None found>"
	}
	Line 0 ""

	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Finished Creating Appendix D - Server Bootstrap Items"
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): "
}
#endregion

#region appendix E
Function OutputAppendixE
{
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Create Appendix E DisableTaskOffload Setting"

	#sort the array by bootstrapname and servername
	$Script:TaskOffloadItems = $Script:TaskOffloadItems | Sort-Object ServerName
	
	If($CSV)
	{
		$File = "$($Script:pwdpath)\$($Script:farm.FarmName)_HealthCheck_AppendixE_DisableTaskOffloadSetting.csv"
		$Script:TaskOffloadItems | Export-CSV -Force -Encoding ASCII -NoTypeInformation -Path $File
	}
	
	Line 0 "Appendix E - DisableTaskOffload Settings"
	Line 0 ""
	Line 0 "Best Practices for Configuring Provisioning Services Server on a Network"
	Line 0 "http://support.citrix.com/article/CTX117374"
	Line 0 "This setting is not needed if you are running PVS 6.0 or later"
	Line 0 ""
	Line 1 "Server Name      DisableTaskOffload Setting" 
	Line 1 "==========================================="
	If($Script:TaskOffloadItems)
	{
		ForEach($Item in $Script:TaskOffloadItems)
		{
			Line 1 ( "{0,-16} {1,-16}" -f $Item.serverName, $Item.TaskOffloadValue )
		}
	}
	Else
	{
		Line 1 "<None found>"
	}
	Line 0 ""

	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Finished Creating Appendix E - DisableTaskOffload Setting"
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): "
}
#endregion

#region appendix F
Function OutputAppendixF
{
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Create Appendix F PVS Services"

	#sort the array by displayname and servername
	$Script:PVSServiceItems = $Script:PVSServiceItems | Sort-Object DisplayName, ServerName
	
	If($CSV)
	{
		#AppendixF and AppendixF2 items are contained in the same array
		$File = "$($Script:pwdpath)\$($Script:farm.FarmName)_HealthCheck_AppendixF_PVSServices.csv"
		$Script:PVSServiceItems | Export-CSV -Force -Encoding ASCII -NoTypeInformation -Path $File
	}
	
	Line 0 "Appendix F - Server PVS Service Items"
	Line 0 ""
	Line 1 "Display Name                      Server Name      Service Name  Status Startup Type Started State   Log on as" 
	Line 1 "========================================================================================================================================"
    ########123456789012345678901234567890123 123456789012345  1234567890123 123456 123456789012 1234567 
	#displayname, servername, name, status, startmode, started, startname, state 
	If($Script:PVSServiceItems)
	{
		ForEach($Item in $Script:PVSServiceItems)
		{
			Line 1 ( "{0,-33} {1,-16} {2,-13} {3,-6} {4,-12} {5,-7} {6,-7} {7,-35}" -f `
			$Item.DisplayName, $Item.serverName, $Item.Name, $Item.Status, $Item.StartMode, `
			$Item.Started, $Item.State, $Item.StartName )
		}
	}
	Else
	{
		Line 1 "<None found>"
	}
	Line 0 ""

	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Finished Creating Appendix F - PVS Services"
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): "
}
#endregion

#region appendix F2
Function OutputAppendixF2
{
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Create Appendix F2 PVS Services Failure Actions"
	#array is already sorted in Function OutputAppendixF
	
	Line 0 "Appendix F2 - Server PVS Service Items Failure Actions"
	Line 0 ""
	Line 1 "Display Name                      Server Name      Service Name  Failure Action 1     Failure Action 2     Failure Action 3    " 
	Line 1 "==============================================================================================================================="
	If($Script:PVSServiceItems)
	{
		ForEach($Item in $Script:PVSServiceItems)
		{
			Line 1 ( "{0,-33} {1,-16} {2,-13} {3,-20} {4,-20} {5,-20}" -f `
			$Item.DisplayName, $Item.serverName, $Item.Name, $Item.FailureAction1, $Item.FailureAction2, $Item.FailureAction3 )
		}
	}
	Else
	{
		Line 1 "<None found>"
	}
	Line 0 ""

	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Finished Creating Appendix F2 - PVS Services Failure Actions"
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): "
}
#endregion

#region appendix G
Function OutputAppendixG
{
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Create Appendix G vDisks to Merge"

	#sort the array
	$Script:VersionsToMerge = $Script:VersionsToMerge | Sort-Object
	
	If($CSV)
	{
		$File = "$($Script:pwdpath)\$($Script:farm.FarmName)_HealthCheck_AppendixG_vDiskstoMerge.csv"
		$Script:VersionsToMerge | Export-CSV -Force -Encoding ASCII -NoTypeInformation -Path $File
	}
	
	Line 0 "Appendix G - vDisks to Consider Merging"
	Line 0 ""
	Line 1 "vDisk Name" 
	Line 1 "========================================"
	If($Script:VersionsToMerge)
	{
		ForEach($Item in $Script:VersionsToMerge)
		{
			Line 1 ( "{0,-40}" -f $Item.vDiskName )
		}
	}
	Else
	{
		Line 1 "<None found>"
	}
	Line 0 ""

	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Finished Creating Appendix G - vDisks to Merge"
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): "
}
#endregion

#region appendix H
Function OutputAppendixH
{
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Create Appendix H Empty Device Collections"

	#sort the array
	$Script:EmptyDeviceCollections = $Script:EmptyDeviceCollections | Sort-Object CollectionName
	
	If($CSV)
	{
		$File = "$($Script:pwdpath)\$($Script:farm.FarmName)_HealthCheck_AppendixH_EmptyDeviceCollections.csv"
		$Script:EmptyDeviceCollections | Export-CSV -Force -Encoding ASCII -NoTypeInformation -Path $File
	}
	
	Line 0 "Appendix H - Empty Device Collections"
	Line 0 ""
	Line 1 "Device Collection Name" 
	Line 1 "=================================================="
	If($Script:EmptyDeviceCollections)
	{
		ForEach($Item in $Script:EmptyDeviceCollections)
		{
			Line 1 ( "{0,-50}" -f $Item.CollectionName )
		}
	}
	Else
	{
		Line 1 "<None found>"
	}
	Line 0 ""

	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Finished Creating Appendix G - Empty Device Collections"
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): "
}
#endregion

#region appendix I 
Function ProcessvDisksWithNoAssociation
{
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Finding vDisks with no Target Device Associations"
	$UnassociatedvDisks = New-Object System.Collections.ArrayList
	$GetWhat = "diskLocator"
	$GetParam = ""
	$ErrorTxt = "Disk Locator information"
	$DiskLocators = BuildPVSObject $GetWhat $GetParam $ErrorTxt
	
	If($Null -eq $DiskLocators)
	{
		Write-Host -foregroundcolor Red -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): No DiskLocators Found"
		OutputAppendixI $Null
	}
	Else
	{
		ForEach($DiskLocator in $DiskLocators)
		{
			#get the diskLocatorId
			$DiskLocatorId = $DiskLocator.diskLocatorId
			
			#now pass the disklocatorid to get device
			#if nothing found, the vDisk is unassociated
			$temp = $DiskLocatorId
			$GetWhat = "device"
			$GetParam = "diskLocatorId = $temp"
			$ErrorTxt = "Device for DiskLocatorId $DiskLocatorId information"
			$Results = BuildPVSObject $GetWhat $GetParam $ErrorTxt
			
			If($Null -ne $Results)
			{
				#device found, vDisk is associated
			}
			Else
			{
				#no device found that uses this vDisk
				$obj1 = [PSCustomObject] @{
					vDiskName = $DiskLocator.diskLocatorName				
				}
				$null = $UnassociatedvDisks.Add($obj1)
			}
		}
		
		If($UnassociatedvDisks.Count -gt 0)
		{
			Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Found $($UnassociatedvDisks.Count) vDisks with no Target Device Associations"
			OutputAppendixI $UnassociatedvDisks
		}
		Else
		{
			Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): All vDisks have Target Device Associations"
			Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): "
			OutputAppendixI $Null
		}
	}
}

Function OutputAppendixI
{
	Param([array]$vDisks)

	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Create Appendix I Unassociated vDisks"

	Line 0 "Appendix I - vDisks with no Target Device Associations"
	Line 0 ""
	Line 1 "vDisk Name" 
	Line 1 "========================================"
	
	If($vDisks)
	{
		#sort the array
		$vDisks = $vDisks | Sort-Object
	
		If($CSV)
		{
			$File = "$($Script:pwdpath)\$($Script:farm.FarmName)_HealthCheck_AppendixI_UnassociatedvDisks.csv"
			$vDisks | Export-CSV -Force -Encoding ASCII -NoTypeInformation -Path $File
		}
	
		ForEach($Item in $vDisks)
		{
			Line 1 ( "{0,-40}" -f $Item.vDiskName )
		}
	}
	Else
	{
		Line 1 "<None found>"
	}
	
	Line 0 ""

	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Finished Creating Appendix I - Unassociated vDisks"
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): "
}
#endregion

#region appendix J
Function OutputAppendixJ
{
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Create Appendix J Bad Streaming IP Addresses"

	#sort the array by bootstrapname and servername
	$Script:BadIPs = $Script:BadIPs | Sort-Object ServerName, IPAddress
	
	If($CSV)
	{
		$File = "$($Script:pwdpath)\$($Script:farm.FarmName)_HealthCheck_AppendixJ_BadStreamingIPAddresses.csv"
		$Script:BadIPs | Export-CSV -Force -Encoding ASCII -NoTypeInformation -Path $File
	}
	
	Line 0 "Appendix J - Bad Streaming IP Addresses"
	Line 0 "Streaming IP addresses that do not exist on the server"
	If($Script:PVSVersion -eq "7")
	{
		Line 0 ""
		Line 1 "Server Name      Streaming IP Address" 
		Line 1 "====================================="
		If($Script:BadIPs) 
		{
			ForEach($Item in $Script:BadIPs)
			{
				Line 1 ( "{0,-16} {1,-16}" -f $Item.serverName, $Item.IPAddress )
			}
		}
		Else
		{
			Line 1 "<None found>"
		}
	}
	Else
	{
		Line 1 "Unable to determine Bad Streaming IP Addresses for PVS versions earlier than 7.0"
	}
	Line 0 ""

	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Finished Creating Appendix J Bad Streaming IP Addresses"
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): "
}
#endregion

#region appendix K
Function OutputAppendixK
{
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Create Appendix K Misc Registry Items"

	#sort the array by regkey, regvalue and servername
	$Script:MiscRegistryItems = $Script:MiscRegistryItems | Sort-Object RegKey, RegValue, ServerName
	
	If($CSV)
	{
		$File = "$($Script:pwdpath)\$($Script:farm.FarmName)_HealthCheck_AppendixK_MiscRegistryItems.csv"
		$Script:MiscRegistryItems | Export-CSV -Force -Encoding ASCII -NoTypeInformation -Path $File
	}
	
	Line 0 "Appendix K - Misc Registry Items"
	Line 0 "Miscellaneous Registry Items That May or May Not Exist on Servers"
	Line 0 "These items may or may not be needed"
	Line 0 "This Appendix is strictly for server comparison only"
	Line 0 ""
	Line 1 "Registry Key                                                                                    Registry Value                                     Data                                                                                       Server Name    " 
	Line 1 "============================================================================================================================================================================================================================================================="
	#       12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345S12345678901234567890123456789012345678901234567890S123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890S123456789012345
	
	$Save = ""
	$First = $True
	If($Script:MiscRegistryItems)
	{
		ForEach($Item in $Script:MiscRegistryItems)
		{
			If(!$First -and $Save -ne "$($Item.RegKey.ToString())$($Item.RegValue.ToString())")
			{
				Line 0 ""
			}

			Line 1 ( "{0,-95} {1,-50} {2,-90} {3,-15}" -f `
			$Item.RegKey, $Item.RegValue, $Item.Value, $Item.serverName )
			$Save = "$($Item.RegKey.ToString())$($Item.RegValue.ToString())"
			If($First)
			{
				$First = $False
			}
		}
	}
	Else
	{
		Line 1 "<None found>"
	}
	Line 0 ""

	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Finished Creating Appendix K Misc Registry Items"
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): "
}
#endregion

#region appendix L
Function OutputAppendixL
{
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Create Appendix L vDisks Configured for Server-Side Caching"
	#sort the array 
	$Script:CacheOnServer = $Script:CacheOnServer | Sort-Object StoreName,SiteName,vDiskName
	
	If($CSV)
	{
		$File = "$($Script:pwdpath)\$($Script:farm.FarmName)_HealthCheck_AppendixL_vDisksConfiguredforServerSideCaching.csv"
		$Script:CacheOnServer | Export-CSV -Force -Encoding ASCII -NoTypeInformation -Path $File
	}

	Line 0 "Appendix L - vDisks Configured for Server Side-Caching"
	Line 0 ""

	If($Script:CacheOnServer)
	{
		Line 1 "Store Name                Site Name                 vDisk Name               "
		Line 1 "============================================================================="
			   #1234567890123456789012345 1234567890123456789012345 1234567890123456789012345

		ForEach($Item in $Script:CacheOnServer)
		{
			Line 1 ( "{0,-25} {1,-25} {2,-25}" -f `
			$Item.StoreName, $Item.SiteName, $Item.vDiskName )
		}
	}
	Else
	{
		Line 1 "<None found>"
	}
	Line 0 ""
	
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Finished Creating Appendix L vDisks Configured for Server-Side Caching"
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): "
}
#endregion

#region appendix M
Function OutputAppendixM
{
	#added in V1.16
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Create Appendix M Microsoft Hotfixes and Updates"

	#sort the array by hotfixid and servername
	$Script:MSHotfixes = $Script:MSHotfixes | Sort-Object HotFixID, ServerName
	
	If($CSV)
	{
		$File = "$($Script:pwdpath)\$($Script:farm.FarmName)_HealthCheck_AppendixM_MicrosoftHotfixesandUpdates.csv"
		$Script:MSHotfixes | Export-CSV -Force -Encoding ASCII -NoTypeInformation -Path $File
	}
	
	Line 0 "Appendix M - Microsoft Hotfixes and Updates"
	Line 0 ""
	Line 1 "Hotfix ID                 Server Name     Caption                                       Description          Installed By                        Installed On Date     "
	Line 1 "======================================================================================================================================================================="
	#       1234567890123456789012345S123456789012345S123456789012345678901234567890123456789012345S12345678901234567890S12345678901234567890123456789012345S1234567890123456789012
	#                                                 http://support.microsoft.com/?kbid=2727528    Security Update      XXX-XX-XDDC01\xxxx.xxxxxx           00/00/0000 00:00:00 PM
	#		25                        15              45                                            20                   35                                  22
	
	$Save = ""
	$First = $True
	If($Script:MSHotfixes)
	{
		ForEach($Item in $Script:MSHotfixes)
		{
			If(!$First -and $Save -ne "$($Item.HotFixID)")
			{
				Line 0 ""
			}

			Line 1 ( "{0,-25} {1,-15} {2,-45} {3,-20} {4,-35} {5,-22}" -f `
			$Item.HotFixID, $Item.ServerName, $Item.Caption, $Item.Description, $Item.InstalledBy, $Item.InstalledOn)
			$Save = "$($Item.HotFixID)"
			If($First)
			{
				$First = $False
			}
		}
	}
	Else
	{
		Line 1 "<None found>"
	}
	Line 0 ""

	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Finished Creating Appendix M Microsoft Hotfixes and Updates"
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): "
}
#endregion

#region appendix N
Function OutputAppendixN
{
	#added in V1.16
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Create Appendix N Windows Installed Components"

	$Script:WinInstalledComponents = $Script:WinInstalledComponents | Sort-Object DisplayName, Name, DDCName
	
	If($CSV)
	{
		$File = "$($Script:pwdpath)\$($Script:farm.FarmName)_HealthCheck_AppendixN_InstalledRolesandFeatures.csv"
		$Script:WinInstalledComponents | Export-CSV -Force -Encoding ASCII -NoTypeInformation -Path $File
	}
	
	Line 0 "Appendix N - Windows Installed Components"
	If($Script:RunningOS -like "*2008*")
	{
		Line 1 "Unable to determine for a Server running Server 2008 or 2008 R2"
		Line 0 ""
	}
	Else
	{
		Line 0 ""
		Line 1 "Display Name                                       Name                           Server Name      Feature Type   "
		Line 1 "=================================================================================================================="
		#       12345678901234567890123456789012345678901234567890S123456789012345678901234567890S123456789012345SS123456789012345
		#       Graphical Management Tools and Infrastructure      NET-Framework-45-Features      XXXXXXXXXXXXXXX  Role Service
		#       50                                                 30                             15               15
		$Save = ""
		$First = $True
		If($Script:WinInstalledComponents)
		{
			ForEach($Item in $Script:WinInstalledComponents)
			{
				If(!$First -and $Save -ne "$($Item.DisplayName)$($Item.Name)")
				{
					Line 0 ""
				}

				Line 1 ( "{0,-50} {1,-30} {2,-15}  {3,-15}" -f `
				$Item.DisplayName, $Item.Name, $Item.ServerName, $Item.FeatureType)
				$Save = "$($Item.DisplayName)$($Item.Name)"
				If($First)
				{
					$First = $False
				}
			}
		}
		Else
		{
			Line 1 "<None found>"
		}
		Line 0 ""
	}

	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Finished Creating Appendix N Windows Installed Components"
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): "
}
#endregion

#region appendix O
Function OutputAppendixO
{
	#added in V1.16
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Create Appendix O PVS Processes"

	$Script:PVSProcessItems = $Script:PVSProcessItems | Sort-Object ProcessName, ServerName
	
	If($CSV)
	{
		$File = "$($Script:pwdpath)\$($Script:farm.FarmName)_HealthCheck_AppendixO_PVSProcesses.csv"
		$Script:PVSProcessItems | Export-CSV -Force -Encoding ASCII -NoTypeInformation -Path $File
	}
	
	Line 0 "Appendix O - PVS Processes"
	Line 0 ""
	Line 1 "Process Name  Server Name     Status     "
	Line 1 "========================================="
	#       1234567890123S123456789012345S12345678901
	#       StreamProcess XXXXXXXXXXXXXXX Not Running
	#       13            15              11
	$Save = ""
	$First = $True
	If($Script:PVSProcessItems)
	{
		ForEach($Item in $Script:PVSProcessItems)
		{
			If(!$First -and $Save -ne "$($Item.ProcessName)")
			{
				Line 0 ""
			}

			Line 1 ( "{0,-13} {1,-15} {2,-11}" -f `
			$Item.ProcessName, $Item.ServerName, $Item.Status)
			$Save = "$($Item.ProcessName)"
			If($First)
			{
				$First = $False
			}
		}
	}
	Else
	{
		Line 1 "<None found>"
	}
	Line 0 ""

	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Finished Creating Appendix O PVS Processes"
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): "
}
#endregion

#region appendix p
Function OutputAppendixP
{
	#added in V1.23
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Create Appendix P Items to Review"

	$Script:ItemsToReview = $Script:ItemsToReview | Sort-Object ItemText
	
	If($CSV)
	{
		$File = "$($Script:pwdpath)\$($Script:farm.FarmName)_HealthCheck_AppendixP_ItemsToReview.csv"
		$Script:ItemsToReview | Export-CSV -Force -Encoding ASCII -NoTypeInformation -Path $File
	}
	
	Line 0 "Appendix P - Items to Review"
	Line 0 ""
	Line 1 "Item                                   "
	Line 1 "======================================="
	#       123456789012345678901234567890134567890
	#       ItemText
	If($Script:ItemsToReview)
	{
		ForEach($Item in $Script:ItemsToReview)
		{
			Line 1 ( "{0,-40}" -f $Item.ItemText)
		}
	}
	Else
	{
		Line 1 "<None found>"
	}
	Line 0 ""

	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Finished Creating Appendix P Items to Review"
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): "
}
#endregion

#region appendix Q
Function OutputAppendixQ
{
	#added in V1.23
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Create Appendix Q Server Items to Review"

	$Script:ServerComputerItemsToReview  = $Script:ServerComputerItemsToReview | Sort-Object ServerName
	$Script:ServerDriveItemsToReview     = $Script:ServerDriveItemsToReview | Sort-Object DriveCaption, ServerName
	$Script:ServerProcessorItemsToReview = $Script:ServerProcessorItemsToReview | Sort-Object ServerName
	$Script:ServerNICItemsToReview       = $Script:ServerNICItemsToReview | Sort-Object ServerName
	
	If($CSV)
	{
		$File = "$($Script:pwdpath)\$($Script:farm.FarmName)_HealthCheck_AppendixQ_ServerComputerItemsToReview.csv"
		$Script:ServerComputerItemsToReview | Export-CSV -Force -Encoding ASCII -NoTypeInformation -Path $File

		$File = "$($Script:pwdpath)\$($Script:farm.FarmName)_HealthCheck_AppendixQ_ServerDriveItemsToReview.csv"
		$Script:ServerDriveItemsToReview | Export-CSV -Force -Encoding ASCII -NoTypeInformation -Path $File

		$File = "$($Script:pwdpath)\$($Script:farm.FarmName)_HealthCheck_AppendixQ_ServerProcessorItemsToReview.csv"
		$Script:ServerProcessorItemsToReview | Export-CSV -Force -Encoding ASCII -NoTypeInformation -Path $File

		$File = "$($Script:pwdpath)\$($Script:farm.FarmName)_HealthCheck_AppendixQ_ServerNICItemsToReview.csv"
		$Script:ServerNICItemsToReview | Export-CSV -Force -Encoding ASCII -NoTypeInformation -Path $File
	}
	
	Line 0 "Appendix Q - Server Items to Review"
	Line 0 ""
	Line 1 "Computer Items to Review"
	Line 2 "Server Name     Operating System                        Power Plan        RAM   Physical  Logical"
	Line 2 "                                                                          (GB)  Procs     Procs  "
	Line 2 "================================================================================================="
	#       123456789012345S12345678901234567890123456789012345678SS1234567890123456SS1234SS12345678SS1234567
	#       XXXXXXXXXXXXXXX Microsoft Windows Server 2019 Standard High performance  9999  999       999
	#       15
	If($Script:ServerComputerItemsToReview)
	{
		ForEach($Item in $Script:ServerComputerItemsToReview)
		{
			Line 2 ( "{0,-15} {1,-38}  {2,-16}  {3,4}  {4,8}  {5,7}" -f `
			$Item.ServerName, $Item.OperatingSystem, $Item.PowerPlan, $Item.TotalRam, $Item.PhysicalProcessors, $Item.LogicalProcessors)
		}
	}
	Else
	{
		Line 2 "<None found>"
	}
	Line 0 ""

	Line 1 "Drive Items to Review"
	Line 2 "Server Name     Caption  Size (GB)"
	Line 2 "=============================================="
	#       123456789012345S1234567SS123456789
	#       XXXXXXXXXXXXXXX C:            9999
	#       15
	If($Script:ServerDriveItemsToReview)
	{
		ForEach($Item in $Script:ServerDriveItemsToReview)
		{
			Line 2 ( "{0,-15} {1,-7}  {2,9}" -f `
			$Item.ServerName, $Item.DriveCaption, $Item.DriveSize)
		}
	}
	Else
	{
		Line 2 "<None found>"
	}
	Line 0 ""

	Line 1 "Processor Items to Review"
	Line 2 "Server Name     Cores  Logical Procs"
	Line 2 "===================================="
	#       123456789012345S12345SS1234567890123
	#       XXXXXXXXXXXXXXX  9999           9999
	#       15
	If($Script:ServerProcessorItemsToReview)
	{
		ForEach($Item in $Script:ServerProcessorItemsToReview)
		{
			Line 2 ( "{0,-15} {1,5}  {2,13}" -f `
			$Item.ServerName, $Item.Cores , $Item.LogicalProcs)
		}
	}
	Else
	{
		Line 2 "<None found>"
	}
	Line 0 ""

	Line 1 "NIC Items to Review"
	Line 2 "Server Name     NIC Name                                    Manufacturer          Power Mgmt  RSS     "
	Line 2 "======================================================================================================"
	#       123456789012345S123456789012345678901234567890123456789012SS12345678901234567890SS1234567890SS12345678
	#       XXXXXXXXXXXXXXX Intel(R) 82574L Gigabit Network Connection  Intel Corporation     Disabled    Disabled
	#       15              42                                          20                    9           8
	If($Script:ServerNICItemsToReview)
	{
		ForEach($Item in $Script:ServerNICItemsToReview)
		{
			Line 2 ( "{0,-15} {1,-42}  {2,-20}  {3,-10}  {4,-8}" -f `
			$Item.ServerName, $Item.Name, $Item.Manufacturer, $Item.PowerMgmt, $Item.RSS)
		}
	}
	Else
	{
		Line 2 "<None found>"
	}
	Line 0 ""

	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Finished Creating Appendix Q Server Items to Review"
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): "
}
#endregion

#region Appendixr
Function OutputAppendixR
{
	#added in V1.24
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Create Appendix R Citrix Installed Components"

	$Script:CtxInstalledComponents = $Script:CtxInstalledComponents | Sort-Object DisplayName, PVSServerName
	
	If($CSV)
	{
		$File = "$($Script:pwdpath)\$($Script:farm.FarmName)_HealthCheck_AppendixR_CitrixInstalledComponents.csv"
		$Script:CtxInstalledComponents | Export-CSV -Force -Encoding ASCII -NoTypeInformation -Path $File *> $Null
	}
	
	Line 0 "Appendix R - Citrix Installed Components"
	Line 0 "This Appendix is for Server comparison only"
	Line 0 ""
	$maxLength = ($Script:CtxInstalledComponents.DisplayName | Measure-Object -Property length -Maximum).Maximum
	$NegativeMaxLength = $maxLength * -1
	Line 1 "Display Name" -nonewline
	Line 0 (" " * ($maxLength - 11)) -nonewline
	Line 0 "Display Version           " -nonewline
	Line 0 "PVS Server Name"
	Line 1 ("=" * ($maxLength + 2 + 15 + 40)) # $maxLength, 2 spaces, "Display Version" plus space, length of Server name
	#Line 1 "Display Name                                                                      Display Version           PVS Server Name                         "
	#Line 1 "====================================================================================================================================================="
	#       123456789012345678901234567890123456789012345678901234567890123456789012345678901SS1234567890123456789012345S1234567890123456789012345678901234567890
	#       Citrix 7.15 LTSR CU4 - Citrix Delegated Administration Service PowerShell snap-in  11.16.6.0 build 33000     DDC123456789012.123456789012345.local 
	#       81                                                                                 25                        40
	
	$Save = ""
	$First = $True
	If($Script:CtxInstalledComponents)
	{
		ForEach($Item in $Script:CtxInstalledComponents)
		{
			If(!$First -and $Save -ne "$($Item.DisplayName)$($Item.DisplayVersion)")
			{
				Line 0 ""
			}

			Line 1 ( "{0,$NegativeMaxLength} {1,-25} {2,-40}" -f `
			$Item.DisplayName, $Item.DisplayVersion, $Item.PVSServerName)
			
			$Save = "$($Item.DisplayName)$($Item.DisplayVersion)"
			If($First)
			{
				$First = $False
			}
		}
	}
	Else
	{
		Line 1 "<None found>"
	}
	Line 0 ""

	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Finished Creating Appendix R Citrix Installed Components"
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): "
}
#endregion

#region save and close document	
Function SaveandCloseTextDocument
{
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Finishing up document"
	#end of document processing

	If( $Host.Version.CompareTo( [System.Version]'2.0' ) -eq 0 )
	{
		Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Saving for PoSH V2"
		Write-Output $global:Output.ToString() | Out-File $Script:Filename1 2>$Null
	}
	Else
	{
		Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Saving for PoSH V3 or later"
		Write-Output $global:Output.ToString() | Out-File $Script:Filename1 4>$Null
	}
}
#endregion

#region script end
Function ProcessScriptEnd
{
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Script has completed"
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): "

	$GotFile = $False

	If(Test-Path "$($Script:FileName1)")
	{
		Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): $($Script:FileName1) is ready for use"
		$GotFile = $True
	}
	Else
	{
		Write-Warning "$(Get-Date -Format G): Unable to save the output file, $($Script:FileName1)"
		Write-Error "Unable to save the output file, $($Script:FileName1)"
	}

	#email output file if requested
	If($GotFile -and ![System.String]::IsNullOrEmpty( $SmtpServer ))
	{
		$emailAttachment = $Script:FileName1

		SendEmail $emailAttachment
	}

	Write-Verbose "$(Get-Date -Format G): Script has completed"
	Write-Verbose "$(Get-Date -Format G): "

	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): "

	#http://poshtips.com/measuring-elapsed-time-in-powershell/
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Script started: $($Script:StartTime)"
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Script ended: $(Get-Date)"
	$runtime = $(Get-Date) - $Script:StartTime
	$Str = [string]::format("{0} days, {1} hours, {2} minutes, {3}.{4} seconds", `
		$runtime.Days, `
		$runtime.Hours, `
		$runtime.Minutes, `
		$runtime.Seconds,
		$runtime.Milliseconds)
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Elapsed time: $($Str)"

	If($Dev)
	{
		If($SmtpServer -eq "")
		{
			Out-File -FilePath $Script:DevErrorFile -InputObject $error 4>$Null
		}
		Else
		{
			Out-File -FilePath $Script:DevErrorFile -InputObject $error -Append 4>$Null
		}
	}

	If($ScriptInfo)
	{
		$SIFile = "$Script:pwdpath\PVSHealthCheckScriptInfo_$(Get-Date -f yyyy-MM-dd_HHmm).txt"
		Out-File -FilePath $SIFile -InputObject "" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "AdminAddress       : $($AdminAddress)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "CSV                : $($CSV)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Dev                : $($Dev)" 4>$Null
		If($Dev)
		{
			Out-File -FilePath $SIFile -Append -InputObject "DevErrorFile       : $($Script:DevErrorFile)" 4>$Null
		}
		Out-File -FilePath $SIFile -Append -InputObject "Domain             : $($Domain)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Filename1          : $($Script:FileName1)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Folder             : $($Folder)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "From               : $($From)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Log                : $($Log)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "PVS Farm Name      : $($Script:farm.farmName)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "PVS Version        : $($Script:PVSFullVersion)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Script Info        : $($ScriptInfo)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Smtp Port          : $($SmtpPort)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Smtp Server        : $($SmtpServer)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Title              : $($Script:Title)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "To                 : $($To)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Use SSL            : $($UseSSL)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "User               : $($User)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "OS Detected        : $($Script:RunningOS)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "PoSH version       : $($Host.Version)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "PSCulture          : $($PSCulture)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "PSUICulture        : $($PSUICulture)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Script start       : $($Script:StartTime)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Elapsed time       : $($Str)" 4>$Null
	}

	#stop transcript logging
	If($Log -eq $True) 
	{
		If($Script:StartLog -eq $true) 
		{
			try 
			{
				Stop-Transcript | Out-Null
				Write-Verbose "$(Get-Date -Format G): $Script:LogPath is ready for use"
			} 
			catch 
			{
				Write-Verbose "$(Get-Date -Format G): Transcript/log stop failed"
			}
		}
	}

	$runtime = $Null
	$Str = $Null
}
#endregion

#region email function
Function SendEmail
{
	Param([array]$Attachments)
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Prepare to email"

	$emailAttachment = $Attachments
	$emailSubject = $Script:Title
	$emailBody = @"
Hello, <br />
<br />
$Script:Title is attached.

"@ 

	If($Dev)
	{
		Out-File -FilePath $Script:DevErrorFile -InputObject $error 4>$Null
	}

	$error.Clear()
	
	If($From -Like "anonymous@*")
	{
		#https://serverfault.com/questions/543052/sending-unauthenticated-mail-through-ms-exchange-with-powershell-windows-server
		$anonUsername = "anonymous"
		$anonPassword = ConvertTo-SecureString -String "anonymous" -AsPlainText -Force
		$anonCredentials = New-Object System.Management.Automation.PSCredential($anonUsername,$anonPassword)

		If($UseSSL)
		{
			Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
			-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To `
			-UseSSL -credential $anonCredentials *>$Null 
		}
		Else
		{
			Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
			-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To `
			-credential $anonCredentials *>$Null 
		}
		
		If($?)
		{
			Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Email successfully sent using anonymous credentials"
		}
		ElseIf(!$?)
		{
			$e = $error[0]

			Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Email was not sent:"
			Write-Warning "$(Get-Date -Format G): Exception: $e.Exception" 
		}
	}
	Else
	{
		If($UseSSL)
		{
			Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Trying to send email using current user's credentials with SSL"
			Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
			-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To `
			-UseSSL *>$Null
		}
		Else
		{
			Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Trying to send email using current user's credentials without SSL"
			Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
			-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To *>$Null
		}

		If(!$?)
		{
			$e = $error[0]
			
			#error 5.7.57 is O365 and error 5.7.0 is gmail
			If($null -ne $e.Exception -and $e.Exception.ToString().Contains("5.7"))
			{
				#The server response was: 5.7.xx SMTP; Client was not authenticated to send anonymous mail during MAIL FROM
				Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Current user's credentials failed. Ask for usable credentials."

				If($Dev)
				{
					Out-File -FilePath $Script:DevErrorFile -InputObject $error -Append 4>$Null
				}

				$error.Clear()

				$emailCredentials = Get-Credential -UserName $From -Message "Enter the password to send email"

				If($UseSSL)
				{
					Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
					-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To `
					-UseSSL -credential $emailCredentials *>$Null 
				}
				Else
				{
					Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
					-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To `
					-credential $emailCredentials *>$Null 
				}

				If($?)
				{
					Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Email successfully sent using new credentials"
				}
				ElseIf(!$?)
				{
					$e = $error[0]

					Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Email was not sent:"
					Write-Warning "$(Get-Date -Format G): Exception: $e.Exception" 
				}
			}
			Else
			{
				Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Email was not sent:"
				Write-Warning "$(Get-Date -Format G): Exception: $e.Exception" 
			}
		}
	}
}
#endregion

Function GetConfigWizardInfo
{
	Param([string]$ComputerName)
	
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): `t`t`t`tGather Config Wizard info"
	$DHCPServicesValue = Get-RegistryValue "HKLM:\SOFTWARE\Citrix\ProvisioningServices\Wizard" "DHCPType" $ComputerName
	$PXEServiceValue = Get-RegistryValue "HKLM:\SOFTWARE\Citrix\ProvisioningServices\Wizard" "PXEType" $ComputerName
	
	$DHCPServices = ""
	$PXEServices = ""

	Switch ($DHCPServicesValue)
	{
		1073741824 {$DHCPServices = "The service that runs on another computer"; Break}
		0 {$DHCPServices = "Microsoft DHCP"; Break}
		1 {$DHCPServices = "Provisioning Services BOOTP service"; Break}
		2 {$DHCPServices = "Other BOOTP or DHCP service"; Break}
		Default {$DHCPServices = "Unable to determine DHCPServices: $($DHCPServicesValue)"; Break}
	}

	If($DHCPServicesValue -eq 1073741824)
	{
		Switch ($PXEServiceValue)
		{
			1073741824 {$PXEServices = "The service that runs on another computer"; Break}
			0 {$PXEServices = "Provisioning Services PXE service"; Break}
			Default {$PXEServices = "Unable to determine PXEServices: $($PXEServiceValue)"; Break}
		}
	}
	ElseIf($DHCPServicesValue -eq 0)
	{
		Switch ($PXEServiceValue)
		{
			1073741824 {$PXEServices = "The service that runs on another computer"; Break}
			0 {$PXEServices = "Microsoft DHCP"; Break}
			1 {$PXEServices = "Provisioning Services PXE service"; Break}
			Default {$PXEServices = "Unable to determine PXEServices: $($PXEServiceValue)"; Break}
		}
	}
	ElseIf($DHCPServicesValue -eq 1)
	{
		$PXEServices = "N/A"
	}
	ElseIf($DHCPServicesValue -eq 2)
	{
		Switch ($PXEServiceValue)
		{
			1073741824 {$PXEServices = "The service that runs on another computer"; Break}
			0 {$PXEServices = "Provisioning Services PXE service"; Break}
			Default {$PXEServices = "Unable to determine PXEServices: $($PXEServiceValue)"; Break}
		}
	}

	$UserAccount1Value = Get-RegistryValue "HKLM:\SOFTWARE\Citrix\ProvisioningServices\Wizard" "Account1" $ComputerName
	$UserAccount3Value = Get-RegistryValue "HKLM:\SOFTWARE\Citrix\ProvisioningServices\Wizard" "Account3" $ComputerName
	
	$UserAccount = ""
	
	If([String]::IsNullOrEmpty($UserAccount1Value) -and $UserAccount3Value -eq 1)
	{
		$UserAccount = "NetWork Service"
	}
	ElseIf([String]::IsNullOrEmpty($UserAccount1Value) -and $UserAccount3Value -eq 0)
	{
		$UserAccount = "Local system account"
	}
	ElseIf(![String]::IsNullOrEmpty($UserAccount1Value))
	{
		$UserAccount = $UserAccount1Value
	}

	$TFTPOptionValue = Get-RegistryValue "HKLM:\SOFTWARE\Citrix\ProvisioningServices\Wizard" "TFTPSetting" $ComputerName
	$TFTPOption = ""
	
	If($TFTPOptionValue -eq 1)
	{
		$TFTPOption = "Yes"
		$TFTPBootstrapLocation = Get-RegistryValue "HKLM:\SOFTWARE\Citrix\ProvisioningServices\Admin" "Bootstrap" $ComputerName
	}
	Else
	{
		$TFTPOption = "No"
	}

	$obj1 = [PSCustomObject] @{
		ServerName        = $ComputerName
		DHCPServicesValue = $DHCPServicesValue
		PXEServicesValue  = $PXEServiceValue
		UserAccount       = $UserAccount
		TFTPOptionValue   = $TFTPOptionValue
	}
	$null = $Script:ConfigWizItems.Add($obj1)
	
	Line 2 "Configuration Wizard Settings"
	Line 3 "DHCP Services`t`t: " $DHCPServices
	Line 3 "PXE Services`t`t: " $PXEServices
	Line 3 "User account`t`t: " $UserAccount
	Line 3 "TFTP Option`t`t: " $TFTPOption
	If($TFTPOptionValue -eq 1)
	{
		Line 3 "TFTP Bootstrap Location`t: " $TFTPBootstrapLocation
	}
	
	Line 0 ""
}

Function GetDisableTaskOffloadInfo
{
	Param([string]$ComputerName)
	
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): `t`t`t`tGather TaskOffload info"
	$TaskOffloadValue = Get-RegistryValue "HKLM:\SYSTEM\CurrentControlSet\Services\TCPIP\Parameters" "DisableTaskOffload" $ComputerName
	
	If($Null -eq $TaskOffloadValue)
	{
		$TaskOffloadValue = "Missing"
	}
	
	$obj1 = [PSCustomObject] @{
		ServerName       = $ComputerName	
		TaskOffloadValue = $TaskOffloadValue	
	}
	$null = $Script:TaskOffloadItems.Add($obj1)
	
	Line 2 "TaskOffload Settings"
	Line 3 "Value: " $TaskOffloadValue
	
	Line 0 ""
}

Function Get-RegKeyToObject 
{
	#function contributed by Andrew Williamson @ Fujitsu Services
    param([string]$RegPath,
    [string]$RegKey,
    [string]$ComputerName)
	
    $val = Get-RegistryValue $RegPath $RegKey $ComputerName
	
    If($Null -eq $val) 
	{
        $tmp = "Not set"
    } 
	Else 
	{
	    $tmp = $val
    }
	
	$obj1 = [PSCustomObject] @{
		ServerName = $ComputerName	
		RegKey     = $RegPath	
		RegValue   = $RegKey	
		Value      = $tmp	
	}
	$null = $Script:MiscRegistryItems.Add($obj1)
}

Function GetMiscRegistryKeys
{
	Param([string]$ComputerName)
	
	#look for the following registry keys and values on PVS servers
		
	#Registry Key                                                      Registry Value                 
	#=================================================================================================
	#HKLM:\SOFTWARE\Citrix\ProvisioningServices                        AutoUpdateUserCache            
	#HKLM:\SOFTWARE\Citrix\ProvisioningServices                        LoggingLevel 
	#HKLM:\SOFTWARE\Citrix\ProvisioningServices                        SkipBootMenu                   
	#HKLM:\SOFTWARE\Citrix\ProvisioningServices                        UseManagementIpInCatalog       
	#HKLM:\SOFTWARE\Citrix\ProvisioningServices                        UseTemplateBootOrder           
	#HKLM:\SOFTWARE\Citrix\ProvisioningServices\IPC                    IPv4Address                    
	#HKLM:\SOFTWARE\Citrix\ProvisioningServices\IPC                    PortBase 
	#HKLM:\SOFTWARE\Citrix\ProvisioningServices\IPC                    PortCount 
	#HKLM:\SOFTWARE\Citrix\ProvisioningServices\Manager                GeneralInetAddr                
	#HKLM:\SOFTWARE\Citrix\ProvisioningServices\MgmtDaemon             IPCTraceFile 
	#HKLM:\SOFTWARE\Citrix\ProvisioningServices\MgmtDaemon             IPCTraceState 
	#HKLM:\SOFTWARE\Citrix\ProvisioningServices\MgmtDaemon             PortOffset 
	#HKLM:\SOFTWARE\Citrix\ProvisioningServices\Notifier               IPCTraceFile 
	#HKLM:\SOFTWARE\Citrix\ProvisioningServices\Notifier               IPCTraceState 
	#HKLM:\SOFTWARE\Citrix\ProvisioningServices\Notifier               PortOffset 
	#HKLM:\SOFTWARE\Citrix\ProvisioningServices\SoapServer             PortOffset 
	#HKLM:\SOFTWARE\Citrix\ProvisioningServices\StreamProcess          IPCTraceFile 
	#HKLM:\SOFTWARE\Citrix\ProvisioningServices\StreamProcess          IPCTraceState 
	#HKLM:\SOFTWARE\Citrix\ProvisioningServices\StreamProcess          PortOffset 
	#HKLM:\SOFTWARE\Citrix\ProvisioningServices\StreamProcess          SkipBootMenu                   
	#HKLM:\SOFTWARE\Citrix\ProvisioningServices\StreamProcess          SkipRIMS                       
	#HKLM:\SOFTWARE\Citrix\ProvisioningServices\StreamProcess          SkipRIMSforPrivate             
	#HKLM:\SYSTEM\CurrentControlSet\Services\BNIStack\Parameters       WcHDNoIntermediateBuffering    
	#HKLM:\SYSTEM\CurrentControlSet\services\BNIStack\Parameters       WcRamConfiguration             
	#HKLM:\SYSTEM\CurrentControlSet\Services\BNIStack\Parameters       WcWarningIncrement             
	#HKLM:\SYSTEM\CurrentControlSet\Services\BNIStack\Parameters       WcWarningPercent               
	#HKLM:\SYSTEM\CurrentControlSet\Services\BNNS\Parameters           EnableOffload                  
	#HKLM:\SYSTEM\Currentcontrolset\services\BNTFTP\Parameters         InitTimeoutSec           
	#HKLM:\SYSTEM\Currentcontrolset\services\BNTFTP\Parameters         MaxBindRetry             
	#HKLM:\SYSTEM\Currentcontrolset\services\PVSTSB\Parameters         InitTimeoutSec           
	#HKLM:\SYSTEM\Currentcontrolset\services\PVSTSB\Parameters         MaxBindRetry      
	
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): `t`t`t`tGather Misc Registry Key data"

	#https://docs.citrix.com/en-us/provisioning/7-1/pvs-readme-7/7-fixed-issues.html
	Get-RegKeyToObject "HKLM:\SOFTWARE\Citrix\ProvisioningServices" "AutoUpdateUserCache" $ComputerName

	#https://support.citrix.com/article/CTX135299
	Get-RegKeyToObject "HKLM:\SOFTWARE\Citrix\ProvisioningServices" "SkipBootMenu" $ComputerName

	#https://support.citrix.com/article/CTX142613
	Get-RegKeyToObject "HKLM:\SOFTWARE\Citrix\ProvisioningServices" "UseManagementIpInCatalog" $ComputerName

	#https://support.citrix.com/article/CTX142613
	Get-RegKeyToObject "HKLM:\SOFTWARE\Citrix\ProvisioningServices" "UseTemplateBootOrder" $ComputerName

	#https://support.citrix.com/article/CTX200196
	Get-RegKeyToObject "HKLM:\SOFTWARE\Citrix\ProvisioningServices\IPC" "UseTemplateBootOrder" $ComputerName

	#https://support.citrix.com/article/CTX200196
	Get-RegKeyToObject "HKLM:\SOFTWARE\Citrix\ProvisioningServices\Manager" "UseTemplateBootOrder" $ComputerName

	#https://support.citrix.com/article/CTX135299
	Get-RegKeyToObject "HKLM:\SOFTWARE\Citrix\ProvisioningServices\StreamProcess" "UseTemplateBootOrder" $ComputerName

	Get-RegKeyToObject "HKLM:\SOFTWARE\Citrix\ProvisioningServices\StreamProcess" "SkipRIMS" $ComputerName

	#https://support.citrix.com/article/CTX200233
	Get-RegKeyToObject "HKLM:\SOFTWARE\Citrix\ProvisioningServices\StreamProcess" "SkipRIMSforPrivate" $ComputerName

	#https://support.citrix.com/article/CTX126042
	Get-RegKeyToObject "HKLM:\SYSTEM\CurrentControlSet\Services\BNIStack\Parameters" "WcHDNoIntermediateBuffering" $ComputerName

	#https://support.citrix.com/article/CTX139849
	Get-RegKeyToObject "HKLM:\SYSTEM\CurrentControlSet\services\BNIStack\Parameters" "WcRamConfiguration" $ComputerName

	#https://docs.citrix.com/en-us/provisioning/7-1/pvs-readme-7/7-fixed-issues.html
	Get-RegKeyToObject "HKLM:\SYSTEM\CurrentControlSet\Services\BNIStack\Parameters" "WcWarningIncrement" $ComputerName

	#https://docs.citrix.com/en-us/provisioning/7-1/pvs-readme-7/7-fixed-issues.html
	Get-RegKeyToObject "HKLM:\SYSTEM\CurrentControlSet\Services\BNIStack\Parameters" "WcWarningPercent" $ComputerName

	#https://support.citrix.com/article/CTX117374
	Get-RegKeyToObject "HKLM:\SYSTEM\CurrentControlSet\Services\BNNS\Parameters" "EnableOffload" $ComputerName
	
	#https://discussions.citrix.com/topic/362671-error-pxe-e53/#entry1863984
	Get-RegKeyToObject "HKLM:\SYSTEM\Currentcontrolset\services\BNTFTP\Parameters" "InitTimeoutSec" $ComputerName
	
	#https://discussions.citrix.com/topic/362671-error-pxe-e53/#entry1863984
	Get-RegKeyToObject "HKLM:\SYSTEM\Currentcontrolset\services\BNTFTP\Parameters" "MaxBindRetry" $ComputerName

	#https://discussions.citrix.com/topic/362671-error-pxe-e53/#entry1863984
	Get-RegKeyToObject "HKLM:\SYSTEM\Currentcontrolset\services\PVSTSB\Parameters" "InitTimeoutSec" $ComputerName
	
	#https://discussions.citrix.com/topic/362671-error-pxe-e53/#entry1863984
	Get-RegKeyToObject "HKLM:\SYSTEM\Currentcontrolset\services\PVSTSB\Parameters" "MaxBindRetry" $ComputerName

	#regkeys recommended by Andrew Williamson @ Fujitsu Services
    Get-RegKeyToObject "HKLM:\SOFTWARE\Citrix\ProvisioningServices" "LoggingLevel" $ComputerName
    Get-RegKeyToObject "HKLM:\SOFTWARE\Citrix\ProvisioningServices\IPC" "PortBase" $ComputerName
    Get-RegKeyToObject "HKLM:\SOFTWARE\Citrix\ProvisioningServices\IPC" "PortCount" $ComputerName
    Get-RegKeyToObject "HKLM:\SOFTWARE\Citrix\ProvisioningServices\MgmtDaemon" "IPCTraceFile" $ComputerName
    Get-RegKeyToObject "HKLM:\SOFTWARE\Citrix\ProvisioningServices\MgmtDaemon" "IPCTraceState" $ComputerName
    Get-RegKeyToObject "HKLM:\SOFTWARE\Citrix\ProvisioningServices\MgmtDaemon" "PortOffset" $ComputerName
    Get-RegKeyToObject "HKLM:\SOFTWARE\Citrix\ProvisioningServices\Notifier" "IPCTraceFile" $ComputerName
    Get-RegKeyToObject "HKLM:\SOFTWARE\Citrix\ProvisioningServices\Notifier" "IPCTraceState" $ComputerName
    Get-RegKeyToObject "HKLM:\SOFTWARE\Citrix\ProvisioningServices\Notifier" "PortOffset" $ComputerName
    Get-RegKeyToObject "HKLM:\SOFTWARE\Citrix\ProvisioningServices\SoapServer" "PortOffset" $ComputerName
    Get-RegKeyToObject "HKLM:\SOFTWARE\Citrix\ProvisioningServices\StreamProcess" "IPCTraceFile" $ComputerName
    Get-RegKeyToObject "HKLM:\SOFTWARE\Citrix\ProvisioningServices\StreamProcess" "IPCTraceState" $ComputerName
    Get-RegKeyToObject "HKLM:\SOFTWARE\Citrix\ProvisioningServices\StreamProcess" "PortOffset" $ComputerName
}

# Gets the specified registry value or $Null if it is missing
Function Get-RegistryValue
{
	[CmdletBinding()]
	Param([string]$path, [string]$name, [string]$ComputerName)
	If($ComputerName -eq $env:computername)
	{
		$key = Get-Item -LiteralPath $path -EA 0
		If($key)
		{
			Return $key.GetValue($name, $Null)
		}
		Else
		{
			Return $Null
		}
	}
	Else
	{
		#path needed here is different for remote registry access
		$path = $path.SubString(6)
		$path2 = $path.Replace('\','\\')
		
		Try
		{
			$Reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $ComputerName)
			$RegKey = $Reg.OpenSubKey($path2)
			If ($RegKey)
			{
				$Results = $RegKey.GetValue($name)

				If($Null -ne $Results)
				{
					Return $Results
				}
				Else
				{
					Return $Null
				}
			}
			Else
			{
				Return $Null
			}
		}
		
		Catch
		{
			Return $Null
		}
	}
}

Function BuildPVSObject
{
	Param([string]$MCLIGetWhat = '', [string]$MCLIGetParameters = '', [string]$TextForErrorMsg = '')

	$error.Clear()

	If($MCLIGetParameters -ne '')
	{
		Try
		{
			$MCLIGetResult = Mcli-Get "$($MCLIGetWhat)" -p "$($MCLIGetParameters)" -EA 0
		}
		
		Catch
		{
			#didn't work
		}
	}
	Else
	{
		Try
		{
			$MCLIGetResult = Mcli-Get "$($MCLIGetWhat)" -EA 0
		}
		
		Catch
		{
			#didn't work
		}
	}

	If($error.Count -eq 0)
	{
		$PluralObject = @()
		$SingleObject = $Null
		ForEach($record in $MCLIGetResult)
		{
			If($record.length -gt 5 -and $record.substring(0,6) -eq "Record")
			{
				If($Null -ne $SingleObject)
				{
					$PluralObject += $SingleObject
				}
				$SingleObject = new-object System.Object
			}

			$index = $record.IndexOf(':')
			If($index -gt 0)
			{
				$property = $record.SubString(0, $index)
				$value    = $record.SubString($index + 2)
				If($property -ne "Executing")
				{
					Add-Member -inputObject $SingleObject -MemberType NoteProperty -Name $property -Value $value
				}
			}
		}
		$PluralObject += $SingleObject
		Return $PluralObject
	}
	Else 
	{
		Line 0 "$($TextForErrorMsg) could not be retrieved"
		Line 0 "Error returned is " $error[0].FullyQualifiedErrorId.Split(',')[0].Trim()
	}
}

Function Check-NeededPSSnapins
{
	Param([parameter(Mandatory = $True)][alias("Snapin")][string[]]$Snapins)

	#Function specifics
	$MissingSnapins = @()
	[bool]$FoundMissingSnapin = $False
	$LoadedSnapins = @()
	$RegisteredSnapins = @()

	#Creates arrays of strings, rather than objects, we're passing strings so this will be more robust.
	$loadedSnapins += Get-PSSnapin | ForEach-Object {$_.name}
	$registeredSnapins += Get-PSSnapin -Registered | ForEach-Object {$_.name}

	ForEach($Snapin in $Snapins)
	{
		#check if the snapin is loaded
		If(!($LoadedSnapins -like $snapin))
		{
			#Check if the snapin is missing
			If(!($RegisteredSnapins -like $Snapin))
			{
				#set the flag if it's not already
				If(!($FoundMissingSnapin))
				{
					$FoundMissingSnapin = $True
				}
				#add the entry to the list
				$MissingSnapins += $Snapin
			}
			Else
			{
				#Snapin is registered, but not loaded, loading it now:
				Write-Host "Loading Windows PowerShell snap-in: $snapin"
				Add-PSSnapin -Name $snapin -EA 0

				If(!($?))
				{
					Write-Error "
	`n`n
	Error loading snapin: $($error[0].Exception.Message)
	`n`n
	Script cannot continue.
	`n`n"
					Return $false
				}				
			}
		}
	}

	If($FoundMissingSnapin)
	{
		Write-Warning "Missing Windows PowerShell snap-ins Detected:"
		$missingSnapins | ForEach-Object {Write-Warning "($_)"}
		Return $False
	}
	Else
	{
		Return $True
	}
}

#region general functions
Function line
#function created by Michael B. Smith, Exchange MVP
#@essentialexch on Twitter
#https://essential.exchange/blog
#for creating the formatted text report
#created March 2011
#updated March 2014
# updated March 2019 to use StringBuilder (about 100 times more efficient than simple strings)
{
	Param
	(
		[Int]    $tabs = 0, 
		[String] $name = '', 
		[String] $value = '', 
		[String] $newline = [System.Environment]::NewLine, 
		[Switch] $nonewline
	)

	while( $tabs -gt 0 )
	{
		#V1.17 - switch to using a StringBuilder for $global:Output
		$null = $global:Output.Append( "`t" )
		$tabs--
	}

	If( $nonewline )
	{
		#V1.17 - switch to using a StringBuilder for $global:Output
		$null = $global:Output.Append( $name + $value )
	}
	Else
	{
		#V1.17 - switch to using a StringBuilder for $global:Output
		$null = $global:Output.AppendLine( $name + $value )
	}
}
#endregion

#script begins

Set-StrictMode -Version Latest

$currentPrincipal = New-Object Security.Principal.WindowsPrincipal( [Security.Principal.WindowsIdentity]::GetCurrent() )

If($currentPrincipal.IsInRole( [Security.Principal.WindowsBuiltInRole]::Administrator ))
{
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): This is an elevated PowerShell session"
}
Else
{
	Write-Error "
	`n`n
	`t`tThis is NOT an elevated PowerShell session.
	`n`n
	`t`tScript will exit.
	`n`n
	"
	Exit
}

If($Folder -ne "")
{
	Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Testing folder path"
	#does it exist
	If(Test-Path $Folder -EA 0)
	{
		#it exists, now check to see if it is a folder and not a file
		If(Test-Path $Folder -pathType Container -EA 0)
		{
			#it exists and it is a folder
			Write-Host -foregroundcolor Yellow -backgroundcolor Black "VERBOSE: $(Get-Date -Format G): Folder path $Folder exists and is a folder"
		}
		Else
		{
			#it exists but it is a file not a folder
			Write-Error "
			`n`n
			`t`t
			Folder $Folder is a file, not a folder.
			`n`n
			`t`t
			Script cannot continue.
			`n`n
			"
			Exit
		}
	}
	Else
	{
		#does not exist
		Write-Error "
		`n`n
		`t`t
		Folder $Folder does not exist.
		`n`n
		`t`t
		Script cannot continue.
		`n`n
		"
		Exit
	}
}

If($Folder -eq "")
{
	$Script:pwdpath = $pwd.Path
}
Else
{
	$Script:pwdpath = $Folder
}

If($Script:pwdpath.EndsWith("\"))
{
	#remove the trailing \
	$Script:pwdpath = $Script:pwdpath.SubString(0, ($Script:pwdpath.Length - 1))
}

#V1.23, add testing for standard Windows folders to keep people from running the script in c:\windows\system32
$BadDir = $False
If($Script:pwdpath -like "*Program*") #should catch Program Files, Program Files (x86), and ProgramData
{
	$BadDir = $True
}
If($Script:pwdpath -like "*PerfLogs*")
{
	$BadDir = $True
}
If($Script:pwdpath -like "*Windows*")
{
	$BadDir = $True
}

#exit script if $BadDir is true
If($BadDir)
{111
	Write-Host "$(Get-Date): 
	
	You are running the script from a standard Windows folder.

	Do not run the script from:

	x:\PerfLogs
	x:\Program Files
	x:\Program Files (x86)
	x:\ProgramData
	x:\Windows or any subfolder

	Script will exit.
	"
	Exit
}

If(![String]::IsNullOrEmpty($SmtpServer) -and [String]::IsNullOrEmpty($From) -and [String]::IsNullOrEmpty($To))
{
	Write-Error "
	`n`n
	`tYou specified an SmtpServer but did not include a From or To email address.
	`n`n
	`tScript cannot continue.
	`n`n"
	Exit
}
If(![String]::IsNullOrEmpty($SmtpServer) -and [String]::IsNullOrEmpty($From) -and ![String]::IsNullOrEmpty($To))
{
	Write-Error "
	`n`n
	`tYou specified an SmtpServer and a To email address but did not include a From email address.
	`n`n
	`tScript cannot continue.
	`n`n"
	Exit
}
If(![String]::IsNullOrEmpty($SmtpServer) -and [String]::IsNullOrEmpty($To) -and ![String]::IsNullOrEmpty($From))
{
	Write-Error "
	`n`n
	`tYou specified an SmtpServer and a From email address but did not include a To email address.
	`n`n
	`tScript cannot continue.
	`n`n"
	Exit
}
If(![String]::IsNullOrEmpty($From) -and ![String]::IsNullOrEmpty($To) -and [String]::IsNullOrEmpty($SmtpServer))
{
	Write-Error "
	`n`n
	`tYou specified From and To email addresses but did not include the SmtpServer.
	`n`n
	`tScript cannot continue.
	`n`n"
	Exit
}
If(![String]::IsNullOrEmpty($From) -and [String]::IsNullOrEmpty($SmtpServer))
{
	Write-Error "
	`n`n
	`tYou specified a From email address but did not include the SmtpServer.
	`n`n
	`tScript cannot continue.
	`n`n"
	Exit
}
If(![String]::IsNullOrEmpty($To) -and [String]::IsNullOrEmpty($SmtpServer))
{
	Write-Error "
	`n`n
	`tYou specified a To email address but did not include the SmtpServer.
	`n`n
	`tScript cannot continue.
	`n`n"
	Exit
}

If($Log) 
{
	#start transcript logging
	$Script:LogPath = "$Script:pwdpath\PVSHealthCheckScriptTranscript_$(Get-Date -f yyyy-MM-dd_HHmm).txt"
	
	try 
	{
		Start-Transcript -Path $Script:LogPath -Force -Verbose:$false | Out-Null
		Write-Verbose "$(Get-Date -Format G): Transcript/log started at $Script:LogPath"
		$Script:StartLog = $true
	} 
	catch 
	{
		Write-Verbose "$(Get-Date -Format G): Transcript/log failed at $Script:LogPath"
		$Script:StartLog = $false
	}
}

If($Dev)
{
	$Error.Clear()
	$Script:DevErrorFile = "$Script:pwdpath\PVSHealthCheckScriptErrors_$(Get-Date -f yyyy-MM-dd_HHmm).txt"
}

[string]$Script:RunningOS = (Get-WmiObject -class Win32_OperatingSystem -EA 0).Caption

$Script:ItemsToReview                = New-Object System.Collections.ArrayList
$Script:ServerComputerItemsToReview  = New-Object System.Collections.ArrayList
$Script:ServerDriveItemsToReview     = New-Object System.Collections.ArrayList
$Script:ServerProcessorItemsToReview = New-Object System.Collections.ArrayList
$Script:ServerNICItemsToReview       = New-Object System.Collections.ArrayList
$Script:AdvancedItems1               = New-Object System.Collections.ArrayList
$Script:AdvancedItems2               = New-Object System.Collections.ArrayList
$Script:ConfigWizItems               = New-Object System.Collections.ArrayList
$Script:BootstrapItems               = New-Object System.Collections.ArrayList
$Script:TaskOffloadItems             = New-Object System.Collections.ArrayList
$Script:PVSServiceItems              = New-Object System.Collections.ArrayList
$Script:VersionsToMerge              = New-Object System.Collections.ArrayList
$Script:NICIPAddresses               = @{}
$Script:StreamingIPAddresses         = New-Object System.Collections.ArrayList
$Script:BadIPs                       = New-Object System.Collections.ArrayList
$Script:EmptyDeviceCollections       = New-Object System.Collections.ArrayList
$Script:MiscRegistryItems            = New-Object System.Collections.ArrayList
$Script:CacheOnServer                = New-Object System.Collections.ArrayList
$Script:MSHotfixes                   = New-Object System.Collections.ArrayList
$Script:WinInstalledComponents       = New-Object System.Collections.ArrayList
$Script:PVSProcessItems              = New-Object System.Collections.ArrayList
$Script:CtxInstalledComponents       = New-Object System.Collections.ArrayList	
$script:startTime                    = Get-Date

# v1.17 - switch to using a StringBuilder for $global:Output
[System.Text.StringBuilder] $global:Output = New-Object System.Text.StringBuilder( 16384 )

CheckOnPoSHPrereqs

SetupRemoting

VerifyPVSServices

GetPVSVersion

GetPVSFarm

ShowScriptOptions

ProcessPVSFarm

ProcessPVSSite

ProcessvDisksinFarm

ProcessStores

OutputAppendixA	#Appendix A - Advanced Server Items (Server/Network)

OutputAppendixB	#Appendix B - Advanced Server Items (Pacing/Device)

OutputAppendixC	#Appendix C - Configuration Wizard Settings

OutputAppendixD	#Appendix D - Server Bootstrap Items

OutputAppendixE	#Appendix E - DisableTaskOffload Settings

OutputAppendixF	#Appendix F - Server PVS Service Items

OutputAppendixF2	#Appendix F2 - Server PVS Service Items Failure Actions

OutputAppendixG	#Appendix G - vDisks to Consider Merging

OutputAppendixH	#Appendix H - Empty Device Collections

#outputs Appendix I - vDisks with no Target Device Associations
ProcessvDisksWithNoAssociation

OutputAppendixJ	#Appendix J - Bad Streaming IP Addresses

OutputAppendixK	#Appendix K - Misc Registry Items

OutputAppendixL	#Appendix L - vDisks Configured for Server Side-Caching

OutputAppendixM	#Appendix M - Microsoft Hotfixes and Updates

OutputAppendixN	#Appendix N - Windows Installed Components

OutputAppendixO	#Appendix O - PVS Processes

OutputAppendixP	#Appendix P - Items to Review

OutputAppendixQ	#Appendix Q - Server Items to Review

OutputAppendixR #Appendix R - Citrix Installed Components

SaveandCloseTextDocument

ProcessScriptEnd
