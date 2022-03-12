#Requires -Version 3.0

<#
.SYNOPSIS
	Creates a basic Health Check of a Citrix PVS 6.x or later farm.
.DESCRIPTION
	Creates a basic health check of a Citrix PVS 6.x or later farm in plain text, 
	HTML, Word, or PDF.

	This script was tested with PVS 6.1 running on Windows Server 2008 R2 and PVS 2112 running 
	on Windows Server 2022.

	Creates a document named after the PVS farm.

	Version 2.0 changes the default output report from text to HTML.

	The script must run from an elevated PowerShell session.

	NOTE: The account used to run this script must have at least Read access to the SQL 
	Server that holds the Citrix Provisioning databases.

	This script is written using the old string-based crappy PowerShell because it 
	supports PVS 6.x.

.PARAMETER AdminAddress
	Specifies the name of a PVS server that the PowerShell script connects to. 
	Using this parameter requires the script to run from an elevated PowerShell session.

	This parameter has an alias of AA
.PARAMETER Domain
	Specifies the domain used for the AdminAddress connection. 

	Default value is contained in $env:UserDomain
.PARAMETER User
	Specifies the user used for the AdminAddress connection. 

	Default value is contained in $env:username
.PARAMETER HTML
	Creates an HTML file with an .html extension.
	
	HTML is now the default report format.
	
	This parameter is set True if no other output format is selected.
.PARAMETER Text
	Creates a formatted text file with a .txt extension.
	
	This parameter is disabled by default.
.PARAMETER CSV
	Creates a CSV file for each Appendix.
	The default value is False.
	
	Output CSV filename is in the format:
	
	PVSFarmName_HealthCheckV2_Appendix#_NameOfAppendix.csv
	
	For example:
		TNPVSFarm_HealthCheckV2_AppendixA_AdvancedServerItems1.csv
		TNPVSFarm_HealthCheckV2_AppendixB_AdvancedServerItems2.csv
		TNPVSFarm_HealthCheckV2_AppendixC_ConfigWizardItems.csv
		TNPVSFarm_HealthCheckV2_AppendixD_ServerBootstrapItems.csv
		TNPVSFarm_HealthCheckV2_AppendixE_DisableTaskOffloadSetting.csv	
		TNPVSFarm_HealthCheckV2_AppendixF_PVSServices.csv
		TNPVSFarm_HealthCheckV2_AppendixG_vDiskstoMerge.csv	
		TNPVSFarm_HealthCheckV2_AppendixH_EmptyDeviceCollections.csv	
		TNPVSFarm_HealthCheckV2_AppendixI_UnassociatedvDisks.csv	
		TNPVSFarm_HealthCheckV2_AppendixJ_BadStreamingIPAddresses.csv	
		TNPVSFarm_HealthCheckV2_AppendixK_MiscRegistryItems.csv
		TNPVSFarm_HealthCheckV2_AppendixL_vDisksConfiguredforServerSideCaching.csv	
		TNPVSFarm_HealthCheckV2_AppendixM_MicrosoftHotfixesandUpdates.csv
		TNPVSFarm_HealthCheckV2_AppendixN_InstalledRolesandFeatures.csv
		TNPVSFarm_HealthCheckV2_AppendixO_PVSProcesses.csv
		TNPVSFarm_HealthCheckV2_AppendixP_ItemsToReview.csv
		TNPVSFarm_HealthCheckV2_AppendixQ_ServerComputerItemsToReview.csv
		TNPVSFarm_HealthCheckV2_AppendixQ_ServerDriveItemsToReview.csv
		TNPVSFarm_HealthCheckV2_AppendixQ_ServerProcessorItemsToReview.csv
		TNPVSFarm_HealthCheckV2_AppendixQ_ServerNICItemsToReview.csv
		TNPVSFarm_HealthCheckV2_AppendixR_CitrixInstalledComponents.csv
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
.PARAMETER AddDateTime
	Adds a date timestamp to the end of the file name.
	
	The timestamp is in the format of yyyy-MM-dd_HHmm.
	June 1, 2022, at 6PM is 2022-06-01_1800.
	
	Output filename will be ReportName_2022-06-01_1800.docx (or .pdf).
	
	This parameter is disabled by default.
	This parameter has an alias of ADT.
.PARAMETER CompanyName
	Company Name to use for the Word Cover Page or the Forest Information section for 
	HTML and Text.
	
	Default value for Word output is contained in 
	HKCU:\Software\Microsoft\Office\Common\UserInfo\CompanyName or
	HKCU:\Software\Microsoft\Office\Common\UserInfo\Company, whichever is populated 
	on the computer running the script.

	This parameter has an alias of CN.

	For Word output, if either registry key does not exist and this parameter is not 
	specified, the report does not contain a Company Name on the cover page.
	
	For HTML and Text output, the Forest Information section does not contain the Company 
	Name if this parameter is not specified.
.PARAMETER ReportFooter
	Outputs a footer section at the end of the report.

	This parameter has an alias of RF.
	
	Report Footer
		Report information:
			Created with: <Script Name> - Release Date: <Script Release Date>
			Script version: <Script Version>
			Started on <Date Time in Local Format>
			Elapsed time: nn days, nn hours, nn minutes, nn.nn seconds
			Ran from domain <Domain Name> by user <Username>
			Ran from the folder <Folder Name>

	Script Name and Script Release date are script-specific variables.
	Start Date Time in Local Format is a script variable.
	Elapsed time is a calculated value.
	Domain Name is $env:USERDNSDOMAIN.
	Username is $env:USERNAME.
	Folder Name is a script variable.
.PARAMETER MSWord
	SaveAs DOCX file
	
	Microsoft Word is no longer the default report format.
	This parameter is disabled by default.
.PARAMETER PDF
	SaveAs PDF file instead of DOCX file.
	
	The PDF file is roughly 5X to 10X larger than the DOCX file.
	
	This parameter requires Microsoft Word to be installed.
	This parameter uses Word's SaveAs PDF capability.

	This parameter is disabled by default.
.PARAMETER CompanyAddress
	Company Address to use for the Cover page if the Cover Page has the Address field.
	
	The following Cover Pages have an Address field:
		Banded (Word 2013/2016)
		Contrast (Word 2010)
		Exposure (Word 2010)
		Filigree (Word 2013/2016)
		Ion (Dark) (Word 2013/2016)
		Retrospect (Word 2013/2016)
		Semaphore (Word 2013/2016)
		Tiles (Word 2010)
		ViewMaster (Word 2013/2016)
		
	This parameter is only valid with the MSWORD and PDF output parameters.
	This parameter has an alias of CA.
.PARAMETER CompanyEmail
	Company Email to use for the Cover page if the Cover Page has the Email field. 
	
	The following Cover Pages have an Email field:
		Facet (Word 2013/2016)
	
	This parameter is only valid with the MSWORD and PDF output parameters.
	This parameter has an alias of CE.
.PARAMETER CompanyFax
	Company Fax to use for the Cover page if the Cover Page has the Fax field. 
	
	The following Cover Pages have a Fax field:
		Contrast (Word 2010)
		Exposure (Word 2010)
	
	This parameter is only valid with the MSWORD and PDF output parameters.
	This parameter has an alias of CF.
.PARAMETER CompanyPhone
	Company Phone to use for the Cover Page if the Cover Page has the Phone field. 
	
	The following Cover Pages have a Phone field:
		Contrast (Word 2010)
		Exposure (Word 2010)
	
	This parameter is only valid with the MSWORD and PDF output parameters.
	This parameter has an alias of CPh.
.PARAMETER CoverPage
	What Microsoft Word Cover Page to use.
	Only Word 2010, 2013, and 2016 are supported.
	(default cover pages in Word en-US)
	
	Valid input is:
		Alphabet (Word 2010. Works)
		Annual (Word 2010. Doesn't work well for this report)
		Austere (Word 2010. Works)
		Austin (Word 2010/2013/2016. Doesn't work in 2013 or 2016, mostly 
		works in 2010 but Subtitle/Subject & Author fields need to be moved 
		after title box is moved up)
		Banded (Word 2013/2016. Works)
		Conservative (Word 2010. Works)
		Contrast (Word 2010. Works)
		Cubicles (Word 2010. Works)
		Exposure (Word 2010. Works if you like looking sideways)
		Facet (Word 2013/2016. Works)
		Filigree (Word 2013/2016. Works)
		Grid (Word 2010/2013/2016. Works in 2010)
		Integral (Word 2013/2016. Works)
		Ion (Dark) (Word 2013/2016. Top date doesn't fit; box needs to be 
		manually resized or font changed to 8 point)
		Ion (Light) (Word 2013/2016. Top date doesn't fit; box needs to be 
		manually resized or font changed to 8 point)
		Mod (Word 2010. Works)
		Motion (Word 2010/2013/2016. Works if top date is manually changed to 
		36 point)
		Newsprint (Word 2010. Works but date is not populated)
		Perspective (Word 2010. Works)
		Pinstripes (Word 2010. Works)
		Puzzle (Word 2010. Top date doesn't fit; box needs to be manually 
		resized or font changed to 14 point)
		Retrospect (Word 2013/2016. Works)
		Semaphore (Word 2013/2016. Works)
		Sideline (Word 2010/2013/2016. Doesn't work in 2013 or 2016, works in 
		2010)
		Slice (Dark) (Word 2013/2016. Doesn't work)
		Slice (Light) (Word 2013/2016. Doesn't work)
		Stacks (Word 2010. Works)
		Tiles (Word 2010. Date doesn't fit unless changed to 26 point)
		Transcend (Word 2010. Works)
		ViewMaster (Word 2013/2016. Works)
		Whisp (Word 2013/2016. Works)
		
	The default value is Sideline.
	This parameter has an alias of CP.
	This parameter is only valid with the MSWORD and PDF output parameters.
.PARAMETER UserName
	Username to use for the Cover Page and Footer.
	Default value is contained in $env:username
	This parameter has an alias of UN.
	This parameter is only valid with the MSWORD and PDF output parameters.
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
	
	Creates an HTML report.
	
	Uses all Default values.
	localhost for AdminAddress.
.EXAMPLE
	PS C:\PSScript > .\PVS_HealthCheck_V2.ps1 -MSWord
	
	Creates a Microsoft Word report.
	
	Uses all Default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the Username.
	localhost for AdminAddress.
.EXAMPLE
	PS C:\PSScript > .\PVS_HealthCheck_V2.ps1 -PDF
	
	Creates an Adobe PDF report.
	
	Uses all Default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the Username.
	localhost for AdminAddress.
.EXAMPLE
	PS C:\PSScript > .\PVS_HealthCheck_V2.ps1 -Text
	
	Creates a plain text report.
	
	Uses all Default values.
	localhost for AdminAddress.
.EXAMPLE
	PS C:\PSScript > .\PVS_HealthCheck_V2.ps1 -AdminAddress PVS1 -User cwebster -Domain 
	WebstersLab

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
	PS C:\PSScript .\PVS_HealthCheck_V2.ps1 -MSWord -CompanyName "Carl Webster 
	Consulting" -CoverPage "Mod" -UserName "Carl Webster" -AdminAddress PVS01

	Creates a Microsoft Word report.
	
	Will use:
		Carl Webster Consulting for the Company Name.
		Mod for the Cover Page format.
		Carl Webster for the Username.

	PVS server named PVS01 for the AdminAddress.
.EXAMPLE
	PS C:\PSScript .\PVS_HealthCheck_V2.ps1 -MSWord -CN "Carl Webster Consulting" 
	-CP "Mod" -UN "Carl Webster" -AA PVS01

	Creates a Microsoft Word report.
	
	Will use:
		Carl Webster Consulting for the Company Name.
		Mod for the Cover Page format.
		Carl Webster for the Username.

	PVS server named PVS01 for the AdminAddress.
.EXAMPLE
	PS C:\PSScript .\PVS_HealthCheck_V2.ps1 -MSWord -CompanyName "Sherlock Holmes 
    Consulting" -CoverPage Exposure -UserName "Dr. Watson" -CompanyAddress "221B Baker 
    Street, London, England" -CompanyFax "+44 1753 276600" -CompanyPhone "+44 1753 276200
	
	Creates a Microsoft Word report.
	
	Will use:
		Sherlock Holmes Consulting for the Company Name.
		Exposure for the Cover Page format.
		Dr. Watson for the Username.
		221B Baker Street, London, England for the Company Address.
		+44 1753 276600 for the Company Fax.
		+44 1753 276200 for the Company Phone.
.EXAMPLE
	PS C:\PSScript > .\PVS_HealthCheck_V2.ps1 -Folder \\FileServer\ShareName
	
	Output file will be saved in the path \\FileServer\ShareName
.EXAMPLE
	PS C:\PSScript > .\PVS_HealthCheck_V2.ps1 -Dev -ScriptInfo -Log
	
	Creates a text file named PVSHealthCheckScriptErrors_yyyyMMddTHHmmssffff.txt that 
	contains up to the last 250 errors reported by the script.
	
	Creates a text file named PVSHealthCheckScriptInfo_yyyy-MM-dd_HHmm.txt that 
	contains all the script parameters and other basic information.
	
	Creates a text file for transcript logging named 
	PVSHealthCheckScriptTranscript_yyyyMMddTHHmmssffff.txt.
.EXAMPLE
	PS C:\PSScript > .\PVS_HealthCheck_V2.ps1 -CSV
	
	Uses all Default values.
	localhost for AdminAddress.
	Creates a CSV file for each Appendix.
.EXAMPLE
	PS C:\PSScript > .\PVS_HealthCheck_V2.ps1 -SmtpServer mail.domain.tld -From 
	ADAdmin@domain.tld -To ITGroup@domain.tld	

	The script uses the email server mail.domain.tld, sending from ADAdmin@domain.tld 
	and sending to ITGroup@domain.tld.

	The script uses the default SMTP port 25 and does not use SSL.

	If the current user's credentials are not valid to send an email, the script prompts 
	the user to enter valid credentials.
.EXAMPLE
	PS C:\PSScript > .\PVS_HealthCheck_v2.ps1 -SmtpServer mailrelay.domain.tld -From 
	Anonymous@domain.tld -To ITGroup@domain.tld	

	***SENDING UNAUTHENTICATED EMAIL***

	The script uses the email server mailrelay.domain.tld, sending from 
	anonymous@domain.tld and sending to ITGroup@domain.tld.

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
	PS C:\PSScript > .\PVS_HealthCheck_v2.ps1 -SmtpServer 
	labaddomain-com.mail.protection.outlook.com -UseSSL -From 
	SomeEmailAddress@labaddomain.com -To ITGroupDL@labaddomain.com	

	***OFFICE 365 Example***

	https://docs.microsoft.com/en-us/exchange/mail-flow-best-practices/how-to-set-up-a-multiFunction-device-or-application-to-send-email-using-office-3
	
	This uses Option 2 from the above link.
	
	***OFFICE 365 Example***

	The script uses the email server labaddomain-com.mail.protection.outlook.com, sending 
	from SomeEmailAddress@labaddomain.com and sending to ITGroupDL@labaddomain.com.

	The script uses the default SMTP port 25 and SSL.
.EXAMPLE
	PS C:\PSScript > .\PVS_HealthCheck_v2.ps1 -SmtpServer smtp.office365.com -SmtpPort 587 
	-UseSSL -From Webster@CarlWebster.com -To ITGroup@CarlWebster.com	

	The script uses the email server smtp.office365.com on port 587 using SSL, sending from 
	webster@carlwebster.com and sending to ITGroup@carlwebster.com.

	If the current user's credentials are not valid to send an email, the script prompts 
	the user to enter valid credentials.
.EXAMPLE
	PS C:\PSScript > .\PVS_HealthCheck_v2.ps1 -SmtpServer smtp.gmail.com -SmtpPort 587 
	-UseSSL -From Webster@CarlWebster.com -To ITGroup@CarlWebster.com	

	*** NOTE ***
	To send an email using a Gmail or g-suite account, you may have to turn ON the "Less 
	secure app access" option on your account.
	*** NOTE ***
	
	The script uses the email server smtp.gmail.com on port 587 using SSL, sending from 
	webster@gmail.com and sending to ITGroup@carlwebster.com.

	If the current user's credentials are not valid to send an email, the script prompts 
	the user to enter valid credentials.
.INPUTS
	None. You cannot pipe objects to this script.
.OUTPUTS
	No objects are output from this script. This script creates a text file and optional 
	CSV files.
.NOTES
	NAME: PVS_HealthCheck_V2.ps1
	VERSION: 2.00
	AUTHOR: Carl Webster (with much help from BG a, now former, Citrix dev)
	LASTEDIT: March 12, 2022
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
	[Switch]$HTML=$False,

	[parameter(Mandatory=$False)] 
	[Switch]$Text=$False,

	[parameter(Mandatory=$False)] 
	[Alias("ADT")]
	[Switch]$AddDateTime=$False,
	
	[parameter(Mandatory=$False)] 
	[Alias("CN")]
	[ValidateNotNullOrEmpty()]
	[string]$CompanyName="",
    
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
	[Alias("RF")]
	[Switch]$ReportFooter=$False,

	[parameter(ParameterSetName="WordPDF",Mandatory=$False)] 
	[Switch]$MSWord=$False,

	[parameter(ParameterSetName="WordPDF",Mandatory=$False)] 
	[Switch]$PDF=$False,

	[parameter(ParameterSetName="WordPDF",Mandatory=$False)] 
	[Alias("CA")]
	[ValidateNotNullOrEmpty()]
	[string]$CompanyAddress="",
    
	[parameter(ParameterSetName="WordPDF",Mandatory=$False)] 
	[Alias("CE")]
	[ValidateNotNullOrEmpty()]
	[string]$CompanyEmail="",
    
	[parameter(ParameterSetName="WordPDF",Mandatory=$False)] 
	[Alias("CF")]
	[ValidateNotNullOrEmpty()]
	[string]$CompanyFax="",
    
	[parameter(ParameterSetName="WordPDF",Mandatory=$False)] 
	[Alias("CPh")]
	[ValidateNotNullOrEmpty()]
	[string]$CompanyPhone="",
    
	[parameter(ParameterSetName="WordPDF",Mandatory=$False)]
	[Alias("CP")]
	[ValidateNotNullOrEmpty()]
	[string]$CoverPage="Sideline", 

	[parameter(ParameterSetName="WordPDF",Mandatory=$False)] 
	[Alias("UN")]
	[ValidateNotNullOrEmpty()]
	[string]$UserName=$env:username,

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
#Version 2.00 14-Mar-2022
#	Added MultiSubnetFailover to Farm Status section
#		Thanks to Arnaud Pain
#		I can't believe no one has asked for this since PVS 7.11 was released on 14-Sep-2016
#	Added the following functions to support Word/PDF and HTML output and new parameters:
#		AddHTMLTable
#		AddWordTable
#		CheckWordPrereq
#		FindWordDocumentEnd
#		FormatHTMLTable
#		GetCulture
#		Get-LocalRegistryValue
#		Get-RegistryValue
#		OutputauthGroups
#		OutputNotice
#		OutputReportFooter
#		OutputWarning
#		ProcessDocumentOutput
#		SaveandCloseDocumentandShutdownWord
#		SaveandCloseHTMLDocument
#		Set-DocumentProperty
#		SetupWord
#		SetWordCellFormat
#		SetWordHashTable
#		SetWordTableAlternateRowColor
#		Test-RegistryValue
#		UpdateDocumentProperties
#		ValidateCompanyName
#		ValidateCoverPage
#		validObject
#		validStateProp
#		WriteHTMLLine
#		WriteWordLine
#	Added the following parameters:
#		AddDateTIme
#		CompanyAddress
#		CompanyEmail
#		CompanyFax
#		CompanyName
#		CompanyPhone
#		CoverPage
#		HTML
#		MSWord
#		PDF
#		ReportFooter
#		Text
#		UserName
#	Any Function in this script not listed anywhere else in this changelog was updated to support Word/PDF and HTML output
#	Changed all file names by adding a V2 somewhere in the name
#	Changed the date format for the transcript and error log files from yyyy-MM-dd_HHmm format to the FileDateTime format
#		The format is yyyyMMddTHHmmssffff (case-sensitive, using a 4-digit year, 2-digit month, 2-digit day, 
#		the letter T as a time separator, 2-digit hour, 2-digit minute, 2-digit second, and 4-digit millisecond)
#		For example: 20221225T0840107271
#	Dropped support for PVS 5. The V1.24 script still supports PVS 5.x
#	Fixed a bug in Function GetInstalledRolesAndFeatures that didn't handle the condition of no installed Roles or Features
#		Thanks to Arnaud Pain for reporting this
#	Fixed a bug when retrieving a Device Collection's Administrators and Operators
#		I was not comparing to the specific device collection name, which returned all administrators and 
#		operators for all device collections and not the device collection being processed 
#	Fixed a bug with handling a PVS server with multiple NICs and multiple IPv4 addresses
#		This bug has existed since I created the script in 2012
#		I changed the $Script:NICIPAddresses array from @{} to New-Object System.Collections.ArrayList
#		Updated Function ProcessPVSSite to use Get-NetIPAddress -CimSession to get all the IPv4 addresses
#		This allows the gathering of IPv4 addresses regardless of the OS or PVS version
#		Updated Function GetBadStreamingIPAddresses to work with the new array type for $Script:NICIPAddresses
#	Format the Farm, Properties, Status section to match the console output
#	In Function DeviceStatus, change the output of the Target device's active/inactive status and license type
#	In Function GetBootstrapInfo, for HTML output, check if any array items exist before outputting a blank table with only column headings
#	In Function GetConfigWizardInfo, fix $PXEServices to work with PVS7+
#		If DHCPType is equal to 1073741824, then if PXE is set to PVS,
#		in PVS V6, PXEType is set to 0, but in PVS7, PXEType is set to 1
#		Updated the function to check for both 0 and 1 values
#	In Function GetDisableTaskOffloadInfo, changed the value of $TaskOffloadValue:
#		If the registry value DisableTaskOffload does not exist, from "Missing" to "Not defined"
#		To a String[] so the HTML output works
#	In Function ProcessStores, change the Word/PDF and HTML output to tables instead of individual lines
#	Replaced all Get-WmiObject with Get-CimInstance
#	Replaced most script Exit calls with AbortScript to stop the transcript log if the log was enabled and started
#	Updated the following functions to the latest versions:
#		AbortScript
#		GetComputerWMIInfo
#		line
#		OutputComputerItem
#		OutputDriveItem
#		OutputNicItem
#		OutputProcessorItem
#		ProcessScriptEnd
#		SaveandCloseTextDocument
#		SendEmail
#		SetFilenames
#		SetupText
#		ShowScriptOptions
#	Updated the help text
#	Updated the ReadMe file
#	Went to Set-StrictMode -Version Latest, from Version 2 and cleaned up all related errors
#	You can select multiple output formats


Function AbortScript
{
	If($MSWord -or $PDF)
	{
		Write-Verbose "$(Get-Date -Format G): System Cleanup"
		If(Test-Path variable:global:word)
		{
			$Script:Word.quit()
			[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Script:Word) | Out-Null
			Remove-Variable -Name word -Scope Global 4>$Null
		}
	}
	[gc]::collect() 
	[gc]::WaitForPendingFinalizers()

	If($MSWord -or $PDF)
	{
		#is the winword Process still running? kill it

		#find out our session (usually "1" except on TS/RDC or Citrix)
		$SessionID = (Get-Process -PID $PID).SessionId

		#Find out if winword running in our session
		$wordprocess = ((Get-Process 'WinWord' -ea 0) | Where-Object {$_.SessionId -eq $SessionID}) | Select-Object -Property Id 
		If( $wordprocess -and $wordprocess.Id -gt 0)
		{
			Write-Verbose "$(Get-Date -Format G): WinWord Process is still running. Attempting to stop WinWord Process # $($wordprocess.Id)"
			Stop-Process $wordprocess.Id -EA 0
		}
	}
	
	Write-Verbose "$(Get-Date -Format G): Script has been aborted"
	#stop transcript logging
	If($Log -eq $True) 
	{
		If($Script:StartLog -eq $True) 
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
	$ErrorActionPreference = $SaveEAPreference
	Exit
}

Set-StrictMode -Version Latest

#force on
$PSDefaultParameterValues = @{"*:Verbose"=$True}
$SaveEAPreference         = $ErrorActionPreference
$ErrorActionPreference    = 'SilentlyContinue'
$global:emailCredentials  = $Null

#Report footer stuff
$script:MyVersion         = 'V2.00'
$Script:ScriptName        = "PVS_HealthCheck_V2.ps1"
$tmpdate                  = [datetime] "03/12/2022"
$Script:ReleaseDate       = $tmpdate.ToUniversalTime().ToShortDateString()

$currentPrincipal = New-Object Security.Principal.WindowsPrincipal( [Security.Principal.WindowsIdentity]::GetCurrent() )

If($currentPrincipal.IsInRole( [Security.Principal.WindowsBuiltInRole]::Administrator ))
{
	Write-Verbose "$(Get-Date -Format G): This is an elevated PowerShell session"
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

If($MSWord -eq $False -and $PDF -eq $False -and $Text -eq $False -and $HTML -eq $False)
{
	$HTML = $True
}

If($MSWord)
{
	Write-Verbose "$(Get-Date -Format G): MSWord is set"
}
If($PDF)
{
	Write-Verbose "$(Get-Date -Format G): PDF is set"
}
If($Text)
{
	Write-Verbose "$(Get-Date -Format G): Text is set"
}
If($HTML)
{
	Write-Verbose "$(Get-Date -Format G): HTML is set"
}

If($Folder -ne "")
{
	Write-Verbose "$(Get-Date -Format G): Testing folder path"
	#does it exist
	If(Test-Path $Folder -EA 0)
	{
		#it exists, now check to see if it is a folder and not a file
		If(Test-Path $Folder -pathType Container -EA 0)
		{
			#it exists and it is a folder
			Write-Verbose "$(Get-Date -Format G): Folder path $Folder exists and is a folder"
		}
		Else
		{
			#it exists but it is a file not a folder
#Do not indent the following write-error lines. Doing so will mess up the console formatting of the error message.
			Write-Error "
			`n`n
	Folder $Folder is a file, not a folder.
			`n`n
	Script cannot continue.
			`n`n"
			AbortScript
		}
	}
	Else
	{
		#does not exist
		Write-Error "
		`n`n
	Folder $Folder does not exist.
		`n`n
	Script cannot continue.
		`n`n
		"
		AbortScript
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

#test for standard Windows folders to keep people from running the script in c:\windows\system32
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
{
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
	AbortScript
}

If(![String]::IsNullOrEmpty($SmtpServer) -and [String]::IsNullOrEmpty($From) -and [String]::IsNullOrEmpty($To))
{
	Write-Error "
	`n`n
	`tYou specified an SmtpServer but did not include a From or To email address.
	`n`n
	`tScript cannot continue.
	`n`n"
	AbortScript
}
If(![String]::IsNullOrEmpty($SmtpServer) -and [String]::IsNullOrEmpty($From) -and ![String]::IsNullOrEmpty($To))
{
	Write-Error "
	`n`n
	`tYou specified an SmtpServer and a To email address but did not include a From email address.
	`n`n
	`tScript cannot continue.
	`n`n"
	AbortScript
}
If(![String]::IsNullOrEmpty($SmtpServer) -and [String]::IsNullOrEmpty($To) -and ![String]::IsNullOrEmpty($From))
{
	Write-Error "
	`n`n
	`tYou specified an SmtpServer and a From email address but did not include a To email address.
	`n`n
	`tScript cannot continue.
	`n`n"
	AbortScript
}
If(![String]::IsNullOrEmpty($From) -and ![String]::IsNullOrEmpty($To) -and [String]::IsNullOrEmpty($SmtpServer))
{
	Write-Error "
	`n`n
	`tYou specified From and To email addresses but did not include the SmtpServer.
	`n`n
	`tScript cannot continue.
	`n`n"
	AbortScript
}
If(![String]::IsNullOrEmpty($From) -and [String]::IsNullOrEmpty($SmtpServer))
{
	Write-Error "
	`n`n
	`tYou specified a From email address but did not include the SmtpServer.
	`n`n
	`tScript cannot continue.
	`n`n"
	AbortScript
}
If(![String]::IsNullOrEmpty($To) -and [String]::IsNullOrEmpty($SmtpServer))
{
	Write-Error "
	`n`n
	`tYou specified a To email address but did not include the SmtpServer.
	`n`n
	`tScript cannot continue.
	`n`n"
	AbortScript
}

If($Log) 
{
	#start transcript logging
	$Script:LogPath = "$Script:pwdpath\PVSHealthCheckScriptTranscript_$(Get-Date -f FileDateTime).txt"
	
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
	$Script:DevErrorFile = "$Script:pwdpath\PVSHealthCheckScriptErrors_$(Get-Date -f FileDateTime).txt"
}

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
$Script:NICIPAddresses               = New-Object System.Collections.ArrayList
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

#region initialize variables for word html and text
[string]$Script:RunningOS = (Get-CIMInstance -ClassName Win32_OperatingSystem -EA 0 -Verbose:$False).Caption
$Script:CoName = $CompanyName #move so this is available for HTML and Text output also

If($MSWord -or $PDF)
{
	#try and fix the issue with the $CompanyName variable
	Write-Verbose "$(Get-Date -Format G): CoName is $($Script:CoName)"
	
	#the following values were attained from 
	#http://msdn.microsoft.com/en-us/library/office/aa211923(v=office.11).aspx
	[int]$wdAlignPageNumberRight = 2
	[int]$wdMove = 0
	[int]$wdSeekMainDocument = 0
	[int]$wdSeekPrimaryFooter = 4
	[int]$wdStory = 6
#	[int]$wdColorBlack = 0
#	[int]$wdColorGray05 = 15987699 
	[int]$wdColorGray15 = 14277081
#	[int]$wdColorRed = 255
	[int]$wdColorWhite = 16777215
#	[int]$wdColorYellow = 65535 #added in ADDS script V2.22
	[int]$wdWord2007 = 12
	[int]$wdWord2010 = 14
	[int]$wdWord2013 = 15
	[int]$wdWord2016 = 16
	[int]$wdFormatDocumentDefault = 16
	[int]$wdFormatPDF = 17
	#http://blogs.technet.com/b/heyscriptingguy/archive/2006/03/01/how-can-i-right-align-a-single-column-in-a-word-table.aspx
	#http://msdn.microsoft.com/en-us/library/office/ff835817%28v=office.15%29.aspx
	#[int]$wdAlignParagraphLeft = 0
	#[int]$wdAlignParagraphCenter = 1
#	[int]$wdAlignParagraphRight = 2
	#http://msdn.microsoft.com/en-us/library/office/ff193345%28v=office.15%29.aspx
#	[int]$wdCellAlignVerticalTop = 0
#	[int]$wdCellAlignVerticalCenter = 1
#	[int]$wdCellAlignVerticalBottom = 2
	#http://msdn.microsoft.com/en-us/library/office/ff844856%28v=office.15%29.aspx
	[int]$wdAutoFitFixed = 0
	[int]$wdAutoFitContent = 1
#	[int]$wdAutoFitWindow = 2
	#http://msdn.microsoft.com/en-us/library/office/ff821928%28v=office.15%29.aspx
	[int]$wdAdjustNone = 0
	[int]$wdAdjustProportional = 1
#	[int]$wdAdjustFirstColumn = 2
#	[int]$wdAdjustSameWidth = 3

	[int]$PointsPerTabStop = 36
	[int]$Indent0TabStops = 0 * $PointsPerTabStop
#	[int]$Indent1TabStops = 1 * $PointsPerTabStop
#	[int]$Indent2TabStops = 2 * $PointsPerTabStop
#	[int]$Indent3TabStops = 3 * $PointsPerTabStop
#	[int]$Indent4TabStops = 4 * $PointsPerTabStop

	#http://www.thedoctools.com/index.php?show=wt_style_names_english_danish_german_french
	[int]$wdStyleHeading1 = -2
	[int]$wdStyleHeading2 = -3
	[int]$wdStyleHeading3 = -4
	[int]$wdStyleHeading4 = -5
	[int]$wdStyleNoSpacing = -158
	[int]$wdTableGrid = -155

	[int]$wdLineStyleNone = 0
	[int]$wdLineStyleSingle = 1

	[int]$wdHeadingFormatTrue = -1
#	[int]$wdHeadingFormatFalse = 0 
}

If($HTML)
{
	#Prior versions used Set-Variable. That hid the variables
	#from @code. So MBS Switched to using $global:

    $global:htmlredmask       = "#FF0000" 4>$Null
    $global:htmlcyanmask      = "#00FFFF" 4>$Null
    $global:htmlbluemask      = "#0000FF" 4>$Null
    $global:htmldarkbluemask  = "#0000A0" 4>$Null
    $global:htmllightbluemask = "#ADD8E6" 4>$Null
    $global:htmlpurplemask    = "#800080" 4>$Null
    $global:htmlyellowmask    = "#FFFF00" 4>$Null
    $global:htmllimemask      = "#00FF00" 4>$Null
    $global:htmlmagentamask   = "#FF00FF" 4>$Null
    $global:htmlwhitemask     = "#FFFFFF" 4>$Null
    $global:htmlsilvermask    = "#C0C0C0" 4>$Null
    $global:htmlgraymask      = "#808080" 4>$Null
    $global:htmlblackmask     = "#000000" 4>$Null
    $global:htmlorangemask    = "#FFA500" 4>$Null
    $global:htmlmaroonmask    = "#800000" 4>$Null
    $global:htmlgreenmask     = "#008000" 4>$Null
    $global:htmlolivemask     = "#808000" 4>$Null

    $global:htmlbold        = 1 4>$Null
    $global:htmlitalics     = 2 4>$Null
	$global:htmlAlignLeft   = 4 4>$Null
	$global:htmlAlignRight  = 8 4>$Null
    $global:htmlred         = 16 4>$Null
    $global:htmlcyan        = 32 4>$Null
    $global:htmlblue        = 64 4>$Null
    $global:htmldarkblue    = 128 4>$Null
    $global:htmllightblue   = 256 4>$Null
    $global:htmlpurple      = 512 4>$Null
    $global:htmlyellow      = 1024 4>$Null
    $global:htmllime        = 2048 4>$Null
    $global:htmlmagenta     = 4096 4>$Null
    $global:htmlwhite       = 8192 4>$Null
    $global:htmlsilver      = 16384 4>$Null
    $global:htmlgray        = 32768 4>$Null
    $global:htmlolive       = 65536 4>$Null
    $global:htmlorange      = 131072 4>$Null
    $global:htmlmaroon      = 262144 4>$Null
    $global:htmlgreen       = 524288 4>$Null
	$global:htmlblack       = 1048576 4>$Null

	$global:htmlsb          = ( $htmlsilver -bor $htmlBold ) ## point optimization

	$global:htmlColor = 
	@{
		$htmlred       = $htmlredmask
		$htmlcyan      = $htmlcyanmask
		$htmlblue      = $htmlbluemask
		$htmldarkblue  = $htmldarkbluemask
		$htmllightblue = $htmllightbluemask
		$htmlpurple    = $htmlpurplemask
		$htmlyellow    = $htmlyellowmask
		$htmllime      = $htmllimemask
		$htmlmagenta   = $htmlmagentamask
		$htmlwhite     = $htmlwhitemask
		$htmlsilver    = $htmlsilvermask
		$htmlgray      = $htmlgraymask
		$htmlolive     = $htmlolivemask
		$htmlorange    = $htmlorangemask
		$htmlmaroon    = $htmlmaroonmask
		$htmlgreen     = $htmlgreenmask
		$htmlblack     = $htmlblackmask
	}
}
#endregion

#region email function
Function SendEmail
{
	Param([array]$Attachments)
	Write-Verbose "$(Get-Date -Format G): Prepare to email"

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
			Write-Verbose "$(Get-Date -Format G): Email successfully sent using anonymous credentials"
		}
		ElseIf(!$?)
		{
			$e = $error[0]

			Write-Verbose "$(Get-Date -Format G): Email was not sent:"
			Write-Warning "$(Get-Date): Exception: $e.Exception" 
		}
	}
	Else
	{
		If($UseSSL)
		{
			Write-Verbose "$(Get-Date -Format G): Trying to send email using current user's credentials with SSL"
			Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
			-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To `
			-UseSSL *>$Null
		}
		Else
		{
			Write-Verbose  "$(Get-Date): Trying to send email using current user's credentials without SSL"
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
				Write-Verbose "$(Get-Date -Format G): Current user's credentials failed. Ask for usable credentials."

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
					Write-Verbose "$(Get-Date -Format G): Email successfully sent using new credentials"
				}
				ElseIf(!$?)
				{
					$e = $error[0]

					Write-Verbose "$(Get-Date -Format G): Email was not sent:"
					Write-Warning "$(Get-Date): Exception: $e.Exception" 
				}
			}
			Else
			{
				Write-Verbose "$(Get-Date -Format G): Email was not sent:"
				Write-Warning "$(Get-Date): Exception: $e.Exception" 
			}
		}
	}
}
#endregion

#region word specific Functions
Function SetWordHashTable
{
	Param([string]$CultureCode)

	#optimized by Michael B. SMith
	
	# DE and FR translations for Word 2010 by Vladimir Radojevic
	# Vladimir.Radojevic@Commerzreal.com

	# DA translations for Word 2010 by Thomas Daugaard
	# Citrix Infrastructure Specialist at edgemo A/S

	# CA translations by Javier Sanchez 
	# CEO & Founder 101 Consulting

	#ca - Catalan
	#da - Danish
	#de - German
	#en - English
	#es - Spanish
	#fi - Finnish
	#fr - French
	#nb - Norwegian
	#nl - Dutch
	#pt - Portuguese
	#sv - Swedish
	#zh - Chinese
	
	[string]$toc = $(
		Switch ($CultureCode)
		{
			'ca-'	{ 'Taula automática 2'; Break }
			'da-'	{ 'Automatisk tabel 2'; Break }
			#'de-'	{ 'Automatische Tabelle 2'; Break }
			'de-'	{ 'Automatisches Verzeichnis 2'; Break } #changed 6-feb-2022 rene bigler
			'de-'	{ 'Automatische Tabelle 2'; Break }
			'en-'	{ 'Automatic Table 2'; Break }
			'es-'	{ 'Tabla automática 2'; Break }
			'fi-'	{ 'Automaattinen taulukko 2'; Break }
			'fr-'	{ 'Table automatique 2'; Break } #changed 13-feb-2017 david roquier and samuel legrand
			'nb-'	{ 'Automatisk tabell 2'; Break }
			'nl-'	{ 'Automatische inhoudsopgave 2'; Break }
			'pt-'	{ 'Sumário Automático 2'; Break }
			'sv-'	{ 'Automatisk innehållsförteckn2'; Break }
			'zh-'	{ '自动目录 2'; Break }
		}
	)

	$Script:myHash                      = @{}
	$Script:myHash.Word_TableOfContents = $toc
	$Script:myHash.Word_NoSpacing       = $wdStyleNoSpacing
	$Script:myHash.Word_Heading1        = $wdStyleheading1
	$Script:myHash.Word_Heading2        = $wdStyleheading2
	$Script:myHash.Word_Heading3        = $wdStyleheading3
	$Script:myHash.Word_Heading4        = $wdStyleheading4
	$Script:myHash.Word_TableGrid       = $wdTableGrid
}

Function GetCulture
{
	Param([int]$WordValue)
	
	#codes obtained from http://msdn.microsoft.com/en-us/library/bb213877(v=office.12).aspx
	$CatalanArray = 1027
	$ChineseArray = 2052,3076,5124,4100
	$DanishArray = 1030
	$DutchArray = 2067, 1043
	$EnglishArray = 3081, 10249, 4105, 9225, 6153, 8201, 5129, 13321, 7177, 11273, 2057, 1033, 12297
	$FinnishArray = 1035
	$FrenchArray = 2060, 1036, 11276, 3084, 12300, 5132, 13324, 6156, 8204, 10252, 7180, 9228, 4108
	$GermanArray = 1031, 3079, 5127, 4103, 2055
	$NorwegianArray = 1044, 2068
	$PortugueseArray = 1046, 2070
	$SpanishArray = 1034, 11274, 16394, 13322, 9226, 5130, 7178, 12298, 17418, 4106, 18442, 19466, 6154, 15370, 10250, 20490, 3082, 14346, 8202
	$SwedishArray = 1053, 2077

	#ca - Catalan
	#da - Danish
	#de - German
	#en - English
	#es - Spanish
	#fi - Finnish
	#fr - French
	#nb - Norwegian
	#nl - Dutch
	#pt - Portuguese
	#sv - Swedish
	#zh - Chinese

	Switch ($WordValue)
	{
		{$CatalanArray -contains $_} {$CultureCode = "ca-"}
		{$ChineseArray -contains $_} {$CultureCode = "zh-"}
		{$DanishArray -contains $_} {$CultureCode = "da-"}
		{$DutchArray -contains $_} {$CultureCode = "nl-"}
		{$EnglishArray -contains $_} {$CultureCode = "en-"}
		{$FinnishArray -contains $_} {$CultureCode = "fi-"}
		{$FrenchArray -contains $_} {$CultureCode = "fr-"}
		{$GermanArray -contains $_} {$CultureCode = "de-"}
		{$NorwegianArray -contains $_} {$CultureCode = "nb-"}
		{$PortugueseArray -contains $_} {$CultureCode = "pt-"}
		{$SpanishArray -contains $_} {$CultureCode = "es-"}
		{$SwedishArray -contains $_} {$CultureCode = "sv-"}
		Default {$CultureCode = "en-"}
	}
	
	Return $CultureCode
}

Function ValidateCoverPage
{
	Param([int]$xWordVersion, [string]$xCP, [string]$CultureCode)
	
	$xArray = ""
	
	Switch ($CultureCode)
	{
		'ca-'	{
				If($xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "En bandes", "Faceta", "Filigrana",
					"Integral", "Ió (clar)", "Ió (fosc)", "Línia lateral",
					"Moviment", "Quadrícula", "Retrospectiu", "Sector (clar)",
					"Sector (fosc)", "Semàfor", "Visualització principal", "Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Austin", "En bandes", "Faceta", "Filigrana",
					"Integral", "Ió (clar)", "Ió (fosc)", "Línia lateral",
					"Moviment", "Quadrícula", "Retrospectiu", "Sector (clar)",
					"Sector (fosc)", "Semàfor", "Visualització", "Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabet", "Anual", "Austin", "Conservador",
					"Contrast", "Cubicles", "Diplomàtic", "Exposició",
					"Línia lateral", "Mod", "Mosiac", "Moviment", "Paper de diari",
					"Perspectiva", "Piles", "Quadrícula", "Sobri",
					"Transcendir", "Trencaclosques")
				}
			}

		'da-'	{
				If($xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "BevægElse", "Brusen", "Facet", "Filigran", 
					"Gitter", "Integral", "Ion (lys)", "Ion (mørk)", 
					"Retro", "Semafor", "Sidelinje", "Stribet", 
					"Udsnit (lys)", "Udsnit (mørk)", "Visningsmaster")
				}
				ElseIf($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("BevægElse", "Brusen", "Ion (lys)", "Filigran",
					"Retro", "Semafor", "Visningsmaster", "Integral",
					"Facet", "Gitter", "Stribet", "Sidelinje", "Udsnit (lys)",
					"Udsnit (mørk)", "Ion (mørk)", "Austin")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("BevægElse", "Moderat", "Perspektiv", "Firkanter",
					"Overskrid", "Alfabet", "Kontrast", "Stakke", "Fliser", "Gåde",
					"Gitter", "Austin", "Eksponering", "Sidelinje", "Enkel",
					"Nålestribet", "Årlig", "Avispapir", "Tradionel")
				}
			}

		'de-'	{
				If($xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "Bewegung", "Facette", "Filigran", 
					"Gebändert", "Integral", "Ion (dunkel)", "Ion (hell)", 
					"Pfiff", "Randlinie", "Raster", "Rückblick", 
					"Segment (dunkel)", "Segment (hell)", "Semaphor", 
					"ViewMaster")
				}
				ElseIf($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Semaphor", "Segment (hell)", "Ion (hell)",
					"Raster", "Ion (dunkel)", "Filigran", "Rückblick", "Pfiff",
					"ViewMaster", "Segment (dunkel)", "Verbunden", "Bewegung",
					"Randlinie", "Austin", "Integral", "Facette")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alphabet", "Austin", "Bewegung", "Durchscheinend",
					"Herausgestellt", "Jährlich", "Kacheln", "Kontrast", "Kubistisch",
					"Modern", "Nadelstreifen", "Perspektive", "Puzzle", "Randlinie",
					"Raster", "Schlicht", "Stapel", "Traditionell", "Zeitungspapier")
				}
			}

		'en-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "Banded", "Facet", "Filigree", "Grid",
					"Integral", "Ion (Dark)", "Ion (Light)", "Motion", "Retrospect",
					"Semaphore", "Sideline", "Slice (Dark)", "Slice (Light)", "ViewMaster",
					"Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alphabet", "Annual", "Austere", "Austin", "Conservative",
					"Contrast", "Cubicles", "Exposure", "Grid", "Mod", "Motion", "Newsprint",
					"Perspective", "Pinstripes", "Puzzle", "Sideline", "Stacks", "Tiles", "Transcend")
				}
			}

		'es-'	{
				If($xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "Con bandas", "Cortar (oscuro)", "Cuadrícula", 
					"Whisp", "Faceta", "Filigrana", "Integral", "Ion (claro)", 
					"Ion (oscuro)", "Línea lateral", "Movimiento", "Retrospectiva", 
					"Semáforo", "Slice (luz)", "Vista principal", "Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Whisp", "Vista principal", "Filigrana", "Austin",
					"Slice (luz)", "Faceta", "Semáforo", "Retrospectiva", "Cuadrícula",
					"Movimiento", "Cortar (oscuro)", "Línea lateral", "Ion (oscuro)",
					"Ion (claro)", "Integral", "Con bandas")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabeto", "Anual", "Austero", "Austin", "Conservador",
					"Contraste", "Cuadrícula", "Cubículos", "Exposición", "Línea lateral",
					"Moderno", "Mosaicos", "Movimiento", "Papel periódico",
					"Perspectiva", "Pilas", "Puzzle", "Rayas", "Sobrepasar")
				}
			}

		'fi-'	{
				If($xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Filigraani", "Integraali", "Ioni (tumma)",
					"Ioni (vaalea)", "Opastin", "Pinta", "Retro", "Sektori (tumma)",
					"Sektori (vaalea)", "Vaihtuvavärinen", "ViewMaster", "Austin",
					"Kuiskaus", "Liike", "Ruudukko", "Sivussa")
				}
				ElseIf($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Filigraani", "Integraali", "Ioni (tumma)",
					"Ioni (vaalea)", "Opastin", "Pinta", "Retro", "Sektori (tumma)",
					"Sektori (vaalea)", "Vaihtuvavärinen", "ViewMaster", "Austin",
					"Kiehkura", "Liike", "Ruudukko", "Sivussa")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Aakkoset", "Askeettinen", "Austin", "Kontrasti",
					"Laatikot", "Liike", "Liituraita", "Mod", "Osittain peitossa",
					"Palapeli", "Perinteinen", "Perspektiivi", "Pinot", "Ruudukko",
					"Ruudut", "Sanomalehtipaperi", "Sivussa", "Vuotuinen", "Ylitys")
				}
			}

		'fr-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("À bandes", "Austin", "Facette", "Filigrane", 
					"Guide", "Intégrale", "Ion (clair)", "Ion (foncé)", 
					"Lignes latérales", "Quadrillage", "Rétrospective", "Secteur (clair)", 
					"Secteur (foncé)", "Sémaphore", "ViewMaster", "Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alphabet", "Annuel", "Austère", "Austin", 
					"Blocs empilés", "Classique", "Contraste", "Emplacements de bureau", 
					"Exposition", "Guide", "Ligne latérale", "Moderne", 
					"Mosaïques", "Mots croisés", "Papier journal", "Perspective",
					"Quadrillage", "Rayures fines", "Transcendant")
				}
			}

		'nb-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "BevegElse", "Dempet", "Fasett", "Filigran",
					"Integral", "Ion (lys)", "Ion (mørk)", "Retrospekt", "Rutenett",
					"Sektor (lys)", "Sektor (mørk)", "Semafor", "Sidelinje", "Stripet",
					"ViewMaster")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabet", "Årlig", "Avistrykk", "Austin", "Avlukker",
					"BevegElse", "Engasjement", "Enkel", "Fliser", "Konservativ",
					"Kontrast", "Mod", "Perspektiv", "Puslespill", "Rutenett", "Sidelinje",
					"Smale striper", "Stabler", "Transcenderende")
				}
			}

		'nl-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "Beweging", "Facet", "Filigraan", "Gestreept",
					"Integraal", "Ion (donker)", "Ion (licht)", "Raster",
					"Segment (Light)", "Semafoor", "Slice (donker)", "Spriet",
					"Terugblik", "Terzijde", "ViewMaster")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Aantrekkelijk", "Alfabet", "Austin", "Bescheiden",
					"Beweging", "Blikvanger", "Contrast", "Eenvoudig", "Jaarlijks",
					"Krantenpapier", "Krijtstreep", "Kubussen", "Mod", "Perspectief",
					"Puzzel", "Raster", "Stapels",
					"Tegels", "Terzijde")
				}
			}

		'pt-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Animação", "Austin", "Em Tiras", "Exibição Mestra",
					"Faceta", "Fatia (Clara)", "Fatia (Escura)", "Filete", "Filigrana", 
					"Grade", "Integral", "Íon (Claro)", "Íon (Escuro)", "Linha Lateral",
					"Retrospectiva", "Semáforo")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabeto", "Animação", "Anual", "Austero", "Austin", "Baias",
					"Conservador", "Contraste", "Exposição", "Grade", "Ladrilhos",
					"Linha Lateral", "Listras", "Mod", "Papel Jornal", "Perspectiva", "Pilhas",
					"Quebra-cabeça", "Transcend")
				}
			}

		'sv-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "Band", "Fasett", "Filigran", "Integrerad", "Jon (ljust)",
					"Jon (mörkt)", "Knippe", "Rutnät", "RörElse", "Sektor (ljus)", "Sektor (mörk)",
					"Semafor", "Sidlinje", "VisaHuvudsida", "Återblick")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabetmönster", "Austin", "Enkelt", "Exponering", "Konservativt",
					"Kontrast", "Kritstreck", "Kuber", "Perspektiv", "Plattor", "Pussel", "Rutnät",
					"RörElse", "Sidlinje", "Sobert", "Staplat", "Tidningspapper", "Årligt",
					"Övergående")
				}
			}

		'zh-'	{
				If($xWordVersion -eq $wdWord2010 -or $xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ('奥斯汀', '边线型', '花丝', '怀旧', '积分',
					'离子(浅色)', '离子(深色)', '母版型', '平面', '切片(浅色)',
					'切片(深色)', '丝状', '网格', '镶边', '信号灯',
					'运动型')
				}
			}

		Default	{
					If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
					{
						$xArray = ("Austin", "Banded", "Facet", "Filigree", "Grid",
						"Integral", "Ion (Dark)", "Ion (Light)", "Motion", "Retrospect",
						"Semaphore", "Sideline", "Slice (Dark)", "Slice (Light)", "ViewMaster",
						"Whisp")
					}
					ElseIf($xWordVersion -eq $wdWord2010)
					{
						$xArray = ("Alphabet", "Annual", "Austere", "Austin", "Conservative",
						"Contrast", "Cubicles", "Exposure", "Grid", "Mod", "Motion", "Newsprint",
						"Perspective", "Pinstripes", "Puzzle", "Sideline", "Stacks", "Tiles", "Transcend")
					}
				}
	}
	
	If($xArray -contains $xCP)
	{
		$xArray = $Null
		Return $True
	}
	Else
	{
		$xArray = $Null
		Return $False
	}
}

Function CheckWordPrereq
{
	If((Test-Path  REGISTRY::HKEY_CLASSES_ROOT\Word.Application) -eq $False)
	{
		$ErrorActionPreference = $SaveEAPreference
		
		If(($MSWord -eq $False) -and ($PDF -eq $True))
		{
			Write-Host "`n`n`t`tThis script uses Microsoft Word's SaveAs PDF function, please install Microsoft Word`n`n"
			AbortScript
		}
		Else
		{
			Write-Host "`n`n`t`tThis script directly outputs to Microsoft Word, please install Microsoft Word`n`n"
			AbortScript
		}
	}

	#find out our session (usually "1" except on TS/RDC or Citrix)
	$SessionID = (Get-Process -PID $PID).SessionId
	
	#Find out if winword is running in our session
	[bool]$wordrunning = $null –ne ((Get-Process 'WinWord' -ea 0) | Where-Object {$_.SessionId -eq $SessionID})
	If($wordrunning)
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Host "`n`n`tPlease close all instances of Microsoft Word before running this report.`n`n"
		AbortScript
	}
}

Function ValidateCompanyName
{
	[bool]$xResult = Test-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "CompanyName"
	If($xResult)
	{
		Return Get-LocalRegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "CompanyName"
	}
	Else
	{
		$xResult = Test-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "Company"
		If($xResult)
		{
			Return Get-LocalRegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "Company"
		}
		Else
		{
			Return ""
		}
	}
}

Function Set-DocumentProperty {
    <#
	.SYNOPSIS
	Function to set the Title Page document properties in MS Word
	.DESCRIPTION
	Long description
	.PARAMETER Document
	Current Document Object
	.PARAMETER DocProperty
	Parameter description
	.PARAMETER Value
	Parameter description
	.EXAMPLE
	Set-DocumentProperty -Document $Script:Doc -DocProperty Title -Value 'MyTitle'
	.EXAMPLE
	Set-DocumentProperty -Document $Script:Doc -DocProperty Company -Value 'MyCompany'
	.EXAMPLE
	Set-DocumentProperty -Document $Script:Doc -DocProperty Author -Value 'Jim Moyle'
	.EXAMPLE
	Set-DocumentProperty -Document $Script:Doc -DocProperty Subject -Value 'MySubjectTitle'
	.NOTES
	Function Created by Jim Moyle June 2017
	Twitter : @JimMoyle
	#>
    param (
        [object]$Document,
        [String]$DocProperty,
        [string]$Value
    )
    try {
        $binding = "System.Reflection.BindingFlags" -as [type]
        $builtInProperties = $Document.BuiltInDocumentProperties
        $property = [System.__ComObject].invokemember("item", $binding::GetProperty, $null, $BuiltinProperties, $DocProperty)
        [System.__ComObject].invokemember("value", $binding::SetProperty, $null, $property, $Value)
    }
    catch {
        Write-Warning "Failed to set $DocProperty to $Value"
    }
}

Function FindWordDocumentEnd
{
	#Return focus to main document    
	$Script:Doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument
	#move to the end of the current document
	$Script:Selection.EndKey($wdStory,$wdMove) | Out-Null
}

Function SetupWord
{
	Write-Verbose "$(Get-Date -Format G): Setting up Word"
    
	If(!$AddDateTime)
	{
		[string]$Script:WordFileName = "$($Script:pwdpath)\$($OutputFileName).docx"
		If($PDF)
		{
			[string]$Script:PDFFileName = "$($Script:pwdpath)\$($OutputFileName).pdf"
		}
	}
	ElseIf($AddDateTime)
	{
		[string]$Script:WordFileName = "$($Script:pwdpath)\$($OutputFileName)_$(Get-Date -f yyyy-MM-dd_HHmm).docx"
		If($PDF)
		{
			[string]$Script:PDFFileName = "$($Script:pwdpath)\$($OutputFileName)_$(Get-Date -f yyyy-MM-dd_HHmm).pdf"
		}
	}

	# Setup word for output
	Write-Verbose "$(Get-Date -Format G): Create Word comObject."
	$Script:Word = New-Object -comobject "Word.Application" -EA 0 4>$Null

#Do not indent the following write-error lines. Doing so will mess up the console formatting of the error message.
	If(!$? -or $Null -eq $Script:Word)
	{
		Write-Warning "The Word object could not be created. You may need to repair your Word installation."
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "
		`n`n
	The Word object could not be created. You may need to repair your Word installation.
		`n`n
	Script cannot Continue.
		`n`n"
		AbortScript
	}

	Write-Verbose "$(Get-Date -Format G): Determine Word language value"
	If( ( validStateProp $Script:Word Language Value__ ) )
	{
		[int]$Script:WordLanguageValue = [int]$Script:Word.Language.Value__
	}
	Else
	{
		[int]$Script:WordLanguageValue = [int]$Script:Word.Language
	}

	If(!($Script:WordLanguageValue -gt -1))
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "
		`n`n
	Unable to determine the Word language value. You may need to repair your Word installation.
		`n`n
	Script cannot Continue.
		`n`n
		"
		AbortScript
	}
	Write-Verbose "$(Get-Date -Format G): Word language value is $($Script:WordLanguageValue)"
	
	$Script:WordCultureCode = GetCulture $Script:WordLanguageValue
	
	SetWordHashTable $Script:WordCultureCode
	
	[int]$Script:WordVersion = [int]$Script:Word.Version
	If($Script:WordVersion -eq $wdWord2016)
	{
		$Script:WordProduct = "Word 2016"
	}
	ElseIf($Script:WordVersion -eq $wdWord2013)
	{
		$Script:WordProduct = "Word 2013"
	}
	ElseIf($Script:WordVersion -eq $wdWord2010)
	{
		$Script:WordProduct = "Word 2010"
	}
	ElseIf($Script:WordVersion -eq $wdWord2007)
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "
		`n`n
	Microsoft Word 2007 is no longer supported.`n`n`t`tScript will end.
		`n`n
		"
		AbortScript
	}
	ElseIf($Script:WordVersion -eq 0)
	{
		Write-Error "
		`n`n
	The Word Version is 0. You should run a full online repair of your Office installation.
		`n`n
	Script cannot Continue.
		`n`n
		"
		AbortScript
	}
	Else
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "
		`n`n
	You are running an untested or unsupported version of Microsoft Word.
		`n`n
	Script will end.
		`n`n
	Please send info on your version of Word to webster@carlwebster.com
		`n`n
		"
		AbortScript
	}

	#only validate CompanyName if the field is blank
	If([String]::IsNullOrEmpty($CompanyName))
	{
		Write-Verbose "$(Get-Date -Format G): Company name is blank. Retrieve company name from registry."
		$TmpName = ValidateCompanyName
		
		If([String]::IsNullOrEmpty($TmpName))
		{
			Write-Host "
		Company Name is blank so Cover Page will not show a Company Name.
		Check HKCU:\Software\Microsoft\Office\Common\UserInfo for Company or CompanyName value.
		You may want to use the -CompanyName parameter if you need a Company Name on the cover page.
			" -ForegroundColor White
			$Script:CoName = $TmpName
		}
		Else
		{
			$Script:CoName = $TmpName
			Write-Verbose "$(Get-Date -Format G): Updated company name to $($Script:CoName)"
		}
	}
	Else
	{
		$Script:CoName = $CompanyName
	}

	If($Script:WordCultureCode -ne "en-")
	{
		Write-Verbose "$(Get-Date -Format G): Check Default Cover Page for $($WordCultureCode)"
		[bool]$CPChanged = $False
		Switch ($Script:WordCultureCode)
		{
			'ca-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Línia lateral"
						$CPChanged = $True
					}
				}

			'da-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Sidelinje"
						$CPChanged = $True
					}
				}

			'de-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Randlinie"
						$CPChanged = $True
					}
				}

			'es-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Línea lateral"
						$CPChanged = $True
					}
				}

			'fi-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Sivussa"
						$CPChanged = $True
					}
				}

			'fr-'	{
					If($CoverPage -eq "Sideline")
					{
						If($Script:WordVersion -eq $wdWord2013 -or $Script:WordVersion -eq $wdWord2016)
						{
							$CoverPage = "Lignes latérales"
							$CPChanged = $True
						}
						Else
						{
							$CoverPage = "Ligne latérale"
							$CPChanged = $True
						}
					}
				}

			'nb-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Sidelinje"
						$CPChanged = $True
					}
				}

			'nl-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Terzijde"
						$CPChanged = $True
					}
				}

			'pt-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Linha Lateral"
						$CPChanged = $True
					}
				}

			'sv-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Sidlinje"
						$CPChanged = $True
					}
				}

			'zh-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "边线型"
						$CPChanged = $True
					}
				}
		}

		If($CPChanged)
		{
			Write-Verbose "$(Get-Date -Format G): Changed Default Cover Page from Sideline to $($CoverPage)"
		}
	}

	Write-Verbose "$(Get-Date -Format G): Validate cover page $($CoverPage) for culture code $($Script:WordCultureCode)"
	[bool]$ValidCP = $False
	
	$ValidCP = ValidateCoverPage $Script:WordVersion $CoverPage $Script:WordCultureCode
	
	If(!$ValidCP)
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Verbose "$(Get-Date -Format G): Word language value $($Script:WordLanguageValue)"
		Write-Verbose "$(Get-Date -Format G): Culture code $($Script:WordCultureCode)"
		Write-Error "
		`n`n
	For $($Script:WordProduct), $($CoverPage) is not a valid Cover Page option.
		`n`n
	Script cannot Continue.
		`n`n
		"
		AbortScript
	}

	$Script:Word.Visible = $False

	#http://jdhitsolutions.com/blog/2012/05/san-diego-2012-powershell-deep-dive-slides-and-demos/
	#using Jeff's Demo-WordReport.ps1 file for examples
	Write-Verbose "$(Get-Date -Format G): Load Word Templates"

	[bool]$Script:CoverPagesExist = $False
	[bool]$BuildingBlocksExist = $False

	$Script:Word.Templates.LoadBuildingBlocks()
	#word 2010/2013/2016
	$BuildingBlocksCollection = $Script:Word.Templates | Where-Object{$_.name -eq "Built-In Building Blocks.dotx"}

	Write-Verbose "$(Get-Date -Format G): Attempt to load cover page $($CoverPage)"
	$part = $Null

	$BuildingBlocksCollection | 
	ForEach-Object {
		If($_.BuildingBlockEntries.Item($CoverPage).Name -eq $CoverPage) 
		{
			$BuildingBlocks = $_
		}
	}        

	If($Null -ne $BuildingBlocks)
	{
		$BuildingBlocksExist = $True

		Try 
		{
			$part = $BuildingBlocks.BuildingBlockEntries.Item($CoverPage)
		}

		Catch
		{
			$part = $Null
		}

		If($Null -ne $part)
		{
			$Script:CoverPagesExist = $True
		}
	}

	If(!$Script:CoverPagesExist)
	{
		Write-Verbose "$(Get-Date -Format G): Cover Pages are not installed or the Cover Page $($CoverPage) does not exist."
		Write-Host "Cover Pages are not installed or the Cover Page $($CoverPage) does not exist." -ForegroundColor White
		Write-Host "This report will not have a Cover Page." -ForegroundColor White
	}

	Write-Verbose "$(Get-Date -Format G): Create empty word doc"
	$Script:Doc = $Script:Word.Documents.Add()
	If($Null -eq $Script:Doc)
	{
		Write-Verbose "$(Get-Date -Format G): "
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "
		`n`n
	An empty Word document could not be created. You may need to repair your Word installation.
		`n`n
	Script cannot Continue.
		`n`n"
		AbortScript
	}

	$Script:Selection = $Script:Word.Selection
	If($Null -eq $Script:Selection)
	{
		Write-Verbose "$(Get-Date -Format G): "
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "
		`n`n
	An unknown error happened selecting the entire Word document for default formatting options.
		`n`n
	Script cannot Continue.
		`n`n"
		AbortScript
	}

	#set Default tab stops to 1/2 inch (this line is not from Jeff Hicks)
	#36 =.50"
	$Script:Word.ActiveDocument.DefaultTabStop = 36

	#Disable Spell and Grammar Check to resolve issue and improve performance (from Pat Coughlin)
	Write-Verbose "$(Get-Date -Format G): Disable grammar and spell checking"
	#bug reported 1-Apr-2014 by Tim Mangan
	#save current options first before turning them off
	$Script:CurrentGrammarOption = $Script:Word.Options.CheckGrammarAsYouType
	$Script:CurrentSpellingOption = $Script:Word.Options.CheckSpellingAsYouType
	$Script:Word.Options.CheckGrammarAsYouType = $False
	$Script:Word.Options.CheckSpellingAsYouType = $False

	If($BuildingBlocksExist)
	{
		#insert new page, getting ready for table of contents
		Write-Verbose "$(Get-Date -Format G): Insert new page, getting ready for table of contents"
		$part.Insert($Script:Selection.Range,$True) | Out-Null
		$Script:Selection.InsertNewPage()

		#table of contents
		Write-Verbose "$(Get-Date -Format G): Table of Contents - $($Script:MyHash.Word_TableOfContents)"
		$toc = $BuildingBlocks.BuildingBlockEntries.Item($Script:MyHash.Word_TableOfContents)
		If($Null -eq $toc)
		{
			Write-Verbose "$(Get-Date -Format G): "
			Write-Host "Table of Content - $($Script:MyHash.Word_TableOfContents) could not be retrieved." -ForegroundColor White
			Write-Host "This report will not have a Table of Contents." -ForegroundColor White
		}
		Else
		{
			$toc.insert($Script:Selection.Range,$True) | Out-Null
		}
	}
	Else
	{
		Write-Host "Table of Contents are not installed." -ForegroundColor White
		Write-Host "Table of Contents are not installed so this report will not have a Table of Contents." -ForegroundColor White
	}

	#set the footer
	Write-Verbose "$(Get-Date -Format G): Set the footer"
	[string]$footertext = "Report created by $username"

	#get the footer
	Write-Verbose "$(Get-Date -Format G): Get the footer and format font"
	$Script:Doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekPrimaryFooter
	#get the footer and format font
	$footers = $Script:Doc.Sections.Last.Footers
	ForEach($footer in $footers) 
	{
		If($footer.exists) 
		{
			$footer.range.Font.name = "Calibri"
			$footer.range.Font.size = 8
			$footer.range.Font.Italic = $True
			$footer.range.Font.Bold = $True
		}
	} #end ForEach
	Write-Verbose "$(Get-Date -Format G): Footer text"
	$Script:Selection.HeaderFooter.Range.Text = $footerText

	#add page numbering
	Write-Verbose "$(Get-Date -Format G): Add page numbering"
	$Script:Selection.HeaderFooter.PageNumbers.Add($wdAlignPageNumberRight) | Out-Null

	FindWordDocumentEnd
	#end of Jeff Hicks 
}

Function UpdateDocumentProperties
{
	Param([string]$AbstractTitle, [string]$SubjectTitle)
	#updated 8-Jun-2017 with additional cover page fields
	#Update document properties
	Write-Verbose "$(Get-Date -Format G): Set Cover Page Properties"
	#8-Jun-2017 put these 4 items in alpha order
	Set-DocumentProperty -Document $Script:Doc -DocProperty Author -Value $UserName
	Set-DocumentProperty -Document $Script:Doc -DocProperty Company -Value $Script:CoName
	Set-DocumentProperty -Document $Script:Doc -DocProperty Subject -Value $SubjectTitle
	Set-DocumentProperty -Document $Script:Doc -DocProperty Title -Value $Script:title

	#Get the Coverpage XML part
	$cp = $Script:Doc.CustomXMLParts | Where-Object{$_.NamespaceURI -match "coverPageProps$"}

	#get the abstract XML part
	$ab = $cp.documentelement.ChildNodes | Where-Object{$_.basename -eq "Abstract"}
	#set the text
	If([String]::IsNullOrEmpty($Script:CoName))
	{
		[string]$abstract = $AbstractTitle
	}
	Else
	{
		[string]$abstract = "$($AbstractTitle) for $($Script:CoName)"
	}
	$ab.Text = $abstract

	#added 8-Jun-2017
	$ab = $cp.documentelement.ChildNodes | Where-Object{$_.basename -eq "CompanyAddress"}
	#set the text
	[string]$abstract = $CompanyAddress
	$ab.Text = $abstract

	#added 8-Jun-2017
	$ab = $cp.documentelement.ChildNodes | Where-Object{$_.basename -eq "CompanyEmail"}
	#set the text
	[string]$abstract = $CompanyEmail
	$ab.Text = $abstract

	#added 8-Jun-2017
	$ab = $cp.documentelement.ChildNodes | Where-Object{$_.basename -eq "CompanyFax"}
	#set the text
	[string]$abstract = $CompanyFax
	$ab.Text = $abstract

	#added 8-Jun-2017
	$ab = $cp.documentelement.ChildNodes | Where-Object{$_.basename -eq "CompanyPhone"}
	#set the text
	[string]$abstract = $CompanyPhone
	$ab.Text = $abstract

	$ab = $cp.documentelement.ChildNodes | Where-Object{$_.basename -eq "PublishDate"}
	#set the text
	[string]$abstract = (Get-Date -Format d).ToString()
	$ab.Text = $abstract

	Write-Verbose "$(Get-Date -Format G): Update the Table of Contents"
	#update the Table of Contents
	$Script:Doc.TablesOfContents.item(1).Update()
	$cp = $Null
	$ab = $Null
	$abstract = $Null
}
#endregion

#region registry Functions
#http://stackoverflow.com/questions/5648931/test-if-registry-value-exists
# This Function just gets $True or $False
Function Test-RegistryValue($path, $name)
{
	$key = Get-Item -LiteralPath $path -EA 0
	$key -and $Null -ne $key.GetValue($name, $Null)
}

# Gets the specified local registry value or $Null if it is missing
Function Get-LocalRegistryValue($path, $name)
{
	$key = Get-Item -LiteralPath $path -EA 0
	If($key)
	{
		$key.GetValue($name, $Null)
	}
	Else
	{
		$Null
	}
}

Function Get-RegistryValue
{
	# Gets the specified registry value or $Null if it is missing
	[CmdletBinding()]
	Param
	(
		[String] $path, 
		[String] $name, 
		[String] $ComputerName
	)

	If($ComputerName -eq $env:computername -or $ComputerName -eq "localhost")
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

	#path needed here is different for remote registry access
	$path1 = $path.SubString( 6 )
	$path2 = $path1.Replace( '\', '\\' )

	$registry = $null
	try
	{
		## use the Remote Registry service
		$registry = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey(
			[Microsoft.Win32.RegistryHive]::LocalMachine,
			$ComputerName ) 
	}
	catch
	{
		#$e = $error[ 0 ]
		#3.06, remove the verbose message as it confised some people
		#wv "Could not open registry on computer $ComputerName ($e)"
	}

	$val = $null
	If( $registry )
	{
		$key = $registry.OpenSubKey( $path2 )
		If( $key )
		{
			$val = $key.GetValue( $name )
			$key.Close()
		}

		$registry.Close()
	}

	Return $val
}
#endregion

#region word, text and html line output Functions
Function line
#Function created by Michael B. Smith, Exchange MVP
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
		#Switch to using a StringBuilder for $global:Output
		$null = $global:Output.Append( "`t" )
		$tabs--
	}

	If( $nonewline )
	{
		#Switch to using a StringBuilder for $global:Output
		$null = $global:Output.Append( $name + $value )
	}
	Else
	{
		#Switch to using a StringBuilder for $global:Output
		$null = $global:Output.AppendLine( $name + $value )
	}
}

Function WriteWordLine
#Function created by Ryan Revord
#@rsrevord on Twitter
#Function created to make output to Word easy in this script
#updated 27-Mar-2014 to include font name, font size, italics and bold options
{
	Param([int]$style=0, 
	[int]$tabs = 0, 
	[string]$name = '', 
	[string]$value = '', 
	[string]$fontName=$Null,
	[int]$fontSize=0,
	[bool]$italics=$False,
	[bool]$boldface=$False,
	[Switch]$nonewline)
	
	#Build output style
	[string]$output = ""
	Switch ($style)
	{
		0 {$Script:Selection.Style = $Script:MyHash.Word_NoSpacing; Break}
		1 {$Script:Selection.Style = $Script:MyHash.Word_Heading1; Break}
		2 {$Script:Selection.Style = $Script:MyHash.Word_Heading2; Break}
		3 {$Script:Selection.Style = $Script:MyHash.Word_Heading3; Break}
		4 {$Script:Selection.Style = $Script:MyHash.Word_Heading4; Break}
		Default {$Script:Selection.Style = $Script:MyHash.Word_NoSpacing; Break}
	}
	
	#build # of tabs
	While($tabs -gt 0)
	{ 
		$output += "`t"; $tabs--; 
	}
 
	If(![String]::IsNullOrEmpty($fontName)) 
	{
		$Script:Selection.Font.name = $fontName
	} 

	If($fontSize -ne 0) 
	{
		$Script:Selection.Font.size = $fontSize
	} 
 
	If($italics -eq $True) 
	{
		$Script:Selection.Font.Italic = $True
	} 
 
	If($boldface -eq $True) 
	{
		$Script:Selection.Font.Bold = $True
	} 

	#output the rest of the parameters.
	$output += $name + $value
	$Script:Selection.TypeText($output)
 
	#test for new WriteWordLine 0.
	If($nonewline)
	{
		# Do nothing.
	} 
	Else 
	{
		$Script:Selection.TypeParagraph()
	}
}

#***********************************************************************************************************
# WriteHTMLLine
#***********************************************************************************************************

<#
.Synopsis
	Writes a line of output for HTML output
.DESCRIPTION
	This Function formats an HTML line
.USAGE
	WriteHTMLLine <Style> <Tabs> <Name> <Value> <Font Name> <Font Size> <Options>

	0 for Font Size denotes using the default font size of 2 or 10 point

.EXAMPLE
	WriteHTMLLine 0 0 " "

	Writes a blank line with no style or tab stops, obviously none needed.

.EXAMPLE
	WriteHTMLLine 0 1 "This is a regular line of text indented 1 tab stops"

	Writes a line with 1 tab stop.

.EXAMPLE
	WriteHTMLLine 0 0 "This is a regular line of text in the default font in italics" "" $Null 0 $htmlitalics

	Writes a line omitting font and font size and setting the italics attribute

.EXAMPLE
	WriteHTMLLine 0 0 "This is a regular line of text in the default font in bold" "" $Null 0 $htmlBold

	Writes a line omitting font and font size and setting the bold attribute

.EXAMPLE
	WriteHTMLLine 0 0 "This is a regular line of text in the default font in bold italics" "" $Null 0 ($htmlBold -bor $htmlitalics)

	Writes a line omitting font and font size and setting both italics and bold options

.EXAMPLE	
	WriteHTMLLine 0 0 "This is a regular line of text in the default font in 10 point" "" $Null 2  # 10 point font

	Writes a line using 10 point font

.EXAMPLE
	WriteHTMLLine 0 0 "This is a regular line of text in Courier New font" "" "Courier New" 0 

	Writes a line using Courier New Font and 0 font point size (default = 2 if set to 0)

.EXAMPLE	
	WriteHTMLLine 0 0 "This is a regular line of RED text indented 0 tab stops with the computer name as data in 10 point Courier New bold italics: " $env:computername "Courier New" 2 ($htmlBold -bor $htmlred -bor $htmlitalics)

	Writes a line using Courier New Font with first and second string values to be used, also uses 10 point font with bold, italics and red color options set.

.NOTES

	Font Size - Unlike word, there is a limited set of font sizes that can be used in HTML. They are:
		0 - default which actually gives it a 2 or 10 point.
		1 - 7.5 point font size
		2 - 10 point
		3 - 13.5 point
		4 - 15 point
		5 - 18 point
		6 - 24 point
		7 - 36 point
	Any number larger than 7 defaults to 7

	Style - Refers to the headers that are used with output and resemble the headers in word, 
	HTML supports headers h1-h6 and h1-h4 are more commonly used. Unlike word, H1 will not 
	give you a blue colored font, you will have to set that yourself.

	Colors and Bold/Italics Flags are:

		htmlbold       
		htmlitalics    
		htmlred        
		htmlcyan        
		htmlblue       
		htmldarkblue   
		htmllightblue   
		htmlpurple      
		htmlyellow      
		htmllime       
		htmlmagenta     
		htmlwhite       
		htmlsilver      
		htmlgray       
		htmlolive       
		htmlorange      
		htmlmaroon      
		htmlgreen       
		htmlblack       
#>

$crlf = [System.Environment]::NewLine

Function WriteHTMLLine
#Function created by Ken Avram
#Function created to make output to HTML easy in this script
#headings fixed 12-Oct-2016 by Webster
#errors with $HTMLStyle fixed 7-Dec-2017 by Webster
# re-implemented/re-based by Michael B. Smith
{
	Param
	(
		[Int]    $style    = 0, 
		[Int]    $tabs     = 0, 
		[String] $name     = '', 
		[String] $value    = '', 
		[String] $fontName = $null,
		[Int]    $fontSize = 1,
		[Int]    $options  = $htmlblack
	)

	[System.Text.StringBuilder] $sb = New-Object System.Text.StringBuilder( 1024 )

	If( [String]::IsNullOrEmpty( $name ) )	
	{
		$null = $sb.Append( '<p></p>' )
	}
	Else
	{
		[Bool] $ital  = $options -band $htmlitalics
		[Bool] $bold  = $options -band $htmlBold
		[Bool] $left  = $options -band $htmlAlignLeft
		[Bool] $right = $options -band $htmlAlignRight

		if( $left )
		{
			$HTMLBody += " align=left"
		}
		elseif( $right )
		{
			$HTMLBody += " align=right"
		}

		If( $ital ) { $null = $sb.Append( '<i>' ) }
		If( $bold ) { $null = $sb.Append( '<b>' ) } 

		Switch( $style )
		{
			1 { $HTMLOpen = '<h1'; $HTMLClose = '</h1>'; Break }
			2 { $HTMLOpen = '<h2'; $HTMLClose = '</h2>'; Break }
			3 { $HTMLOpen = '<h3'; $HTMLClose = '</h3>'; Break }
			4 { $HTMLOpen = '<h4'; $HTMLClose = '</h4>'; Break }
			Default { $HTMLOpen = ''; $HTMLClose = ''; Break }
		}

		$null = $sb.Append( $HTMLOpen )
		if( $HTMLOpen.Length -gt 0 )
		{
			if( $left )
			{
				$null = $sb.Append( ' align=left' )
			}
			elseif( $right )
			{
				$null = $sb.Append( ' align=right' )
			}
			$null = $sb.Append( '>' )
		}

		$null = $sb.Append( ( '&nbsp;&nbsp;&nbsp;&nbsp;' * $tabs ) + $name + $value )

		If( $HTMLClose -eq '' ) { $null = $sb.Append( '<br>' )     }
		Else                    { $null = $sb.Append( $HTMLClose ) }

		If( $ital ) { $null = $sb.Append( '</i>' ) }
		If( $bold ) { $null = $sb.Append( '</b>' ) } 

		If( $HTMLClose -eq '' ) { $null = $sb.Append( '<br />' ) }
	}
	$null = $sb.AppendLine( '' )

	Out-File -FilePath $Script:HTMLFileName -Append -InputObject $sb.ToString() 4>$Null
}
#endregion

#region HTML table Functions
#***********************************************************************************************************
# AddHTMLTable - Called from FormatHTMLTable Function
# Created by Ken Avram
# modified by Jake Rutski
# re-implemented by Michael B. Smith. Also made the documentation match reality.
#***********************************************************************************************************
Function AddHTMLTable
{
	Param
	(
		[String]   $fontName  = 'Calibri',
		[Int]      $fontSize  = 2,
		[Int]      $colCount  = 0,
		[Int]      $rowCount  = 0,
		[Object[]] $rowInfo   = $null,
		[Object[]] $fixedInfo = $null
	)

	$fwLength = if( $null -ne $fixedInfo ) { $fixedInfo.Count } else { 0 }

	[System.Text.StringBuilder] $sb = New-Object System.Text.StringBuilder( 8192 )

	If( $rowInfo -and $rowInfo.Length -lt $rowCount )
	{
		$rowCount = $rowInfo.Length
	}

	for( $rowCountIndex = 0; $rowCountIndex -lt $rowCount; $rowCountIndex++ )
	{
		$null = $sb.AppendLine( '<tr>' )

		## reset
		$row = $rowInfo[ $rowCountIndex ]

		$subRow = $row
		If( $subRow -is [Array] -and $subRow[ 0 ] -is [Array] )
		{
			$subRow = $subRow[ 0 ]
		}

		$subRowLength = $subRow.Count
		for( $columnIndex = 0; $columnIndex -lt $colCount; $columnIndex += 2 )
		{
			$item = If( $columnIndex -lt $subRowLength ) { $subRow[ $columnIndex ] } Else { 0 }

			$text   = If( $item ) { $item.ToString() } Else { '' }
			$format = If( ( $columnIndex + 1 ) -lt $subRowLength ) { $subRow[ $columnIndex + 1 ] } Else { 0 }
			## item, text, and format ALWAYS have values, even if empty values
			$color  = $global:htmlColor[ $format -band 0xfffff0 ]
			[Bool] $bold  = $format -band $htmlBold
			[Bool] $ital  = $format -band $htmlitalics
			[Bool] $left  = $format -band $htmlAlignLeft
			[Bool] $right = $format -band $htmlAlignRight		

			If( $fwLength -eq 0 )
			{
				$null = $sb.Append( "<td style=""background-color:$( $color )""" )
			}
			Else
			{
				$null = $sb.Append( "<td style=""width:$( $fixedInfo[ $columnIndex / 2 ] ); background-color:$( $color )""" )
			}
			if( $left )
			{
				$null = $sb.Append( ' align=left' )
			}
			elseif( $right )
			{
				$null = $sb.Append( ' align=right' )
			}
			$null = $sb.Append( "><font face='$($fontName)' size='$($fontSize)'>" )

			If( $bold ) { $null = $sb.Append( '<b>' ) }
			If( $ital ) { $null = $sb.Append( '<i>' ) }

			If( $text -eq ' ' -or $text.length -eq 0)
			{
				$null = $sb.Append( '&nbsp;&nbsp;&nbsp;' )
			}
			Else
			{
				for ($inx = 0; $inx -lt $text.length; $inx++ )
				{
					If( $text[ $inx ] -eq ' ' )
					{
						$null = $sb.Append( '&nbsp;' )
					}
					Else
					{
						break
					}
				}
				$null = $sb.Append( $text )
			}

			If( $bold ) { $null = $sb.Append( '</b>' ) }
			If( $ital ) { $null = $sb.Append( '</i>' ) }

			$null = $sb.AppendLine( '</font></td>' )
		}

		$null = $sb.AppendLine( '</tr>' )
	}

	Out-File -FilePath $Script:HTMLFileName -Append -InputObject $sb.ToString() 4>$Null 
}

#***********************************************************************************************************
# FormatHTMLTable 
# Created by Ken Avram
# modified by Jake Rutski
# reworked by Michael B. Smith
#***********************************************************************************************************

<#
.Synopsis
	Formats table column headers for an HTML table.
.DESCRIPTION
	This function formats table column headers for an HTML table. It requires 
	AddHTMLTable to format the individual rows of the table.

.PARAMETER noBorder
	If set to $true, a table will be generated without a border (border = '0'). 
	Otherwise the table will be generated with a border (border = '1').

.PARAMETER noHeadCols
	This parameter should be used when generating tables which do not have a 
	separate array containing column headers (columnArray is not specified). 

	Set this parameter equal to the number of (header) columns in the table.

.PARAMETER rowArray
	This parameter contains the row data array for the table.

	The total numbers of rows in the table is equal to $rowArray.Length + $tableHeader.Length.
	$tableHeader.Length may be zero (the parameter can be $null).

	Each entry in rowarray is ANOTHER array of tuples. The first element of the tuple is the 
	contents of the cell, and the second element of the tuple is the color of the cell, then
	they duplicate for every cell in the row.

.PARAMETER columnArray
	This parameter contains column header data for the table.

	The total number of columns in the table is equal to $columnarray.Length or $null.

	If $columnarray is $null, then there are no column headers, just the first line of the
	table and noHeadCols is used to size the table.

	Each entry in $columnarray organized as a set of two items. The first is the
	data for the header cell. THe second is the color/italic/bold for the header cell. So
	the total number of columns is ($columnArray.Length / 2) when $columnArray isn't $null.

	I have no idea why it wasn't done identically to $rowarray.

.PARAMETER fixedWidth
	This parameter contains widths for columns in pixel format ("100px") to override auto column widths
	The variable should contain a width for each column you wish to override the auto-size setting
	For example: $fixedWidth = @("100px","110px","120px","130px","140px")

	This is mapped to both rowArray and columnArray.

.PARAMETER tableHeader
	A string containing the header for the table (printed at the top of the table, left justified). The
	default is a blank string.
.PARAMETER tableWidth
	The width of the table in pixels, or 'auto'. The default is 'auto'.
.PARAMETER fontName
	The name of the font to use in the table. The default is 'Calibri'.
.PARAMETER fontSize
	The size of the font to use in the table. The default is 2. Note that this is the HTML size, not the pixel size.

.USAGE
	FormatHTMLTable <Table Header> <Table Width> <Font Name> <Font Size>

.EXAMPLE
	FormatHTMLTable "Table Heading" "auto" "Calibri" 3

	This example formats a table and writes it out into an html file. All of the parameters are optional
	defaults are used if not supplied.

	for <Table format>, the default is auto which will autofit the text into the columns and adjust to the longest text in that column. You can also use percentage i.e. 25%
	which will take only 25% of the line and will auto word wrap the text to the next line in the column. Also, instead of using a percentage, you can use pixels i.e. 400px.

	FormatHTMLTable "Table Heading" "auto" -rowArray $rowData -columnArray $columnData

	This example creates an HTML table with a heading of 'Table Heading', auto column spacing, column header data from $columnData and row data from $rowData

	FormatHTMLTable "Table Heading" -rowArray $rowData -noHeadCols 3

	This example creates an HTML table with a heading of 'Table Heading', auto column spacing, no header, and row data from $rowData

	FormatHTMLTable "Table Heading" -rowArray $rowData -fixedWidth $fixedColumns

	This example creates an HTML table with a heading of 'Table Heading, no header, row data from $rowData, and fixed columns defined by $fixedColumns

.NOTES
	In order to use the formatted table it first has to be loaded with data. Examples below will show how to load the table:

	First, initialize the table array

	$rowdata = @()

	Then Load the array. If you are using column headers then load those into the column headers array, otherwise the first line of the table goes into the column headers array
	and the second and subsequent lines go into the $rowdata table as shown below:

	$columnHeaders = @('Display Name',$htmlsb,'Status',$htmlsb,'Startup Type',$htmlsb)

	The first column is the actual name to display, the second are the attributes of the column i.e. color anded with bold or italics. For the anding, parens are required or it will
	not format correctly.

	This is following by adding rowdata as shown below. As more columns are added the columns will auto adjust to fit the size of the page.

	$rowdata = @()
	$columnHeaders = @("User Name",$htmlsb,$UserName,$htmlwhite)
	$rowdata += @(,('Save as PDF',$htmlsb,$PDF.ToString(),$htmlwhite))
	$rowdata += @(,('Save as TEXT',$htmlsb,$TEXT.ToString(),$htmlwhite))
	$rowdata += @(,('Save as WORD',$htmlsb,$MSWORD.ToString(),$htmlwhite))
	$rowdata += @(,('Save as HTML',$htmlsb,$HTML.ToString(),$htmlwhite))
	$rowdata += @(,('Add DateTime',$htmlsb,$AddDateTime.ToString(),$htmlwhite))
	$rowdata += @(,('Hardware Inventory',$htmlsb,$Hardware.ToString(),$htmlwhite))
	$rowdata += @(,('Computer Name',$htmlsb,$ComputerName,$htmlwhite))
	$rowdata += @(,('Filename1',$htmlsb,$Script:FileName1,$htmlwhite))
	$rowdata += @(,('OS Detected',$htmlsb,$Script:RunningOS,$htmlwhite))
	$rowdata += @(,('PSUICulture',$htmlsb,$PSCulture,$htmlwhite))
	$rowdata += @(,('PoSH version',$htmlsb,$Host.Version.ToString(),$htmlwhite))
	FormatHTMLTable "Example of Horizontal AutoFitContents HTML Table" -rowArray $rowdata

	The 'rowArray' paramater is mandatory to build the table, but it is not set as such in the Function - if nothing is passed, the table will be empty.

	Colors and Bold/Italics Flags are shown below:

		htmlbold       
		htmlitalics    
		htmlred        
		htmlcyan        
		htmlblue       
		htmldarkblue   
		htmllightblue   
		htmlpurple      
		htmlyellow      
		htmllime       
		htmlmagenta     
		htmlwhite       
		htmlsilver      
		htmlgray       
		htmlolive       
		htmlorange      
		htmlmaroon      
		htmlgreen       
		htmlblack     

#>

Function FormatHTMLTable
{
	Param
	(
		[String]   $tableheader = '',
		[String]   $tablewidth  = 'auto',
		[String]   $fontName    = 'Calibri',
		[Int]      $fontSize    = 2,
		[Switch]   $noBorder    = $false,
		[Int]      $noHeadCols  = 1,
		[Object[]] $rowArray    = $null,
		[Object[]] $fixedWidth  = $null,
		[Object[]] $columnArray = $null
	)

	$HTMLBody = ''
	if( $tableheader.Length -gt 0 )
	{
		$HTMLBody += "<b><font face='" + $fontname + "' size='" + ($fontsize + 1) + "'>" + $tableheader + "</font></b>" + $crlf
	}

	$fwSize = if( $null -eq $fixedWidth ) { 0 } else { $fixedWidth.Count }

	If( $null -eq $columnArray -or $columnArray.Length -eq 0)
	{
		$NumCols = $noHeadCols + 1
	}  # means we have no column headers, just a table
	Else
	{
		$NumCols = $columnArray.Length
	}  # need to add one for the color attrib

	If( $null -eq $rowArray )
	{
		$NumRows = 1
	}
	Else
	{
		$NumRows = $rowArray.length + 1
	}

	If( $noBorder )
	{
		$HTMLBody += "<table border='0' width='" + $tablewidth + "'>"
	}
	Else
	{
		$HTMLBody += "<table border='1' width='" + $tablewidth + "'>"
	}
	$HTMLBody += $crlf

	If( $columnArray -and $columnArray.Length -gt 0 )
	{
		$HTMLBody += '<tr>' + $crlf

		for( $columnIndex = 0; $columnIndex -lt $NumCols; $columnindex += 2 )
		{
			$val = $columnArray[ $columnIndex + 1 ]
			$tmp = $global:htmlColor[ $val -band 0xfffff0 ]
			[Bool] $bold  = $val -band $htmlBold
			[Bool] $ital  = $val -band $htmlitalics
			[Bool] $left  = $val -band $htmlAlignLeft
			[Bool] $right = $val -band $htmlAlignRight		

			If( $fwSize -eq 0 )
			{
				$HTMLBody += "<td style=""background-color:$($tmp)"""
			}
			Else
			{
				$HTMLBody += "<td style=""width:$($fixedWidth[$columnIndex / 2]); background-color:$($tmp)"""
			}
			if( $left )
			{
				$HTMLBody += " align=left"
			}
			elseif( $right )
			{
				$HTMLBody += " align=right"
			}

			$HTMLBody += "><font face='$($fontName)' size='$($fontSize)'>"

			If( $bold ) { $HTMLBody += '<b>' }
			If( $ital ) { $HTMLBody += '<i>' }

			$array = $columnArray[ $columnIndex ]
			If( $array )
			{
				If( $array -eq ' ' -or $array.Length -eq 0 )
				{
					$HTMLBody += '&nbsp;&nbsp;&nbsp;'
				}
				Else
				{
					for( $i = 0; $i -lt $array.Length; $i += 2 )
					{
						If( $array[ $i ] -eq ' ' )
						{
							$HTMLBody += '&nbsp;'
						}
						Else
						{
							break
						}
					}
					$HTMLBody += $array
				}
			}
			Else
			{
				$HTMLBody += '&nbsp;&nbsp;&nbsp;'
			}
			
			If( $bold ) { $HTMLBody += '</b>' }
			If( $ital ) { $HTMLBody += '</i>' }

			$HTMLBody += '</font></td>'
			$HTMLBody += $crlf
		}

		$HTMLBody += '</tr>' + $crlf
	}

	Out-File -FilePath $Script:HTMLFileName -Append -InputObject $HTMLBody 4>$Null 
	$HTMLBody = ''

	If( $rowArray )
	{

		AddHTMLTable -fontName $fontName -fontSize $fontSize `
			-colCount $numCols -rowCount $NumRows `
			-rowInfo $rowArray -fixedInfo $fixedWidth
		$rowArray = $null
		$HTMLBody = '</table>'
	}
	Else
	{
		$HTMLBody += '</table>'
	}

	Out-File -FilePath $Script:HTMLFileName -Append -InputObject $HTMLBody 4>$Null 
}
#endregion

#region other HTML Functions
Function SetupHTML
{
	Write-Verbose "$(Get-Date -Format G): Setting up HTML"
	If(!$AddDateTime)
	{
		[string]$Script:HTMLFileName = "$($Script:pwdpath)\$($OutputFileName).html"
	}
	ElseIf($AddDateTime)
	{
		[string]$Script:HTMLFileName = "$($Script:pwdpath)\$($OutputFileName)_$(Get-Date -f yyyy-MM-dd_HHmm).html"
	}

	$htmlhead = "<html><head><meta http-equiv='Content-Language' content='da'><title>" + $Script:Title + "</title></head><body>"
	out-file -FilePath $Script:HTMLFileName -Force -InputObject $HTMLHead 4>$Null
}
#endregion

#region Iain's Word table Functions

<#
.Synopsis
	Add a table to a Microsoft Word document
.DESCRIPTION
	This Function adds a table to a Microsoft Word document from either an array of
	Hashtables or an array of PSCustomObjects.

	Using this Function is quicker than setting each table cell individually but can
	only utilise the built-in MS Word table autoformats. Individual tables cells can
	be altered after the table has been appended to the document (a table reference
	is Returned).
.EXAMPLE
	AddWordTable -Hashtable $HashtableArray

	This example adds table to the MS Word document, utilising all key/value pairs in
	the array of hashtables. Column headers will display the key names as defined.
	Note: the columns might not be displayed in the order that they were defined. To
	ensure columns are displayed in the required order utilise the -Columns parameter.
.EXAMPLE
	AddWordTable -Hashtable $HashtableArray -List

	This example adds table to the MS Word document, utilising all key/value pairs in
	the array of hashtables. No column headers will be added, in a ListView format.
	Note: the columns might not be displayed in the order that they were defined. To
	ensure columns are displayed in the required order utilise the -Columns parameter.
.EXAMPLE
	AddWordTable -CustomObject $PSCustomObjectArray

	This example adds table to the MS Word document, utilising all note property names
	the array of PSCustomObjects. Column headers will display the note property names.
	Note: the columns might not be displayed in the order that they were defined. To
	ensure columns are displayed in the required order utilise the -Columns parameter.
.EXAMPLE
	AddWordTable -Hashtable $HashtableArray -Columns FirstName,LastName,EmailAddress

	This example adds a table to the MS Word document, but only using the specified
	key names: FirstName, LastName and EmailAddress. If other keys are present in the
	array of Hashtables they will be ignored.
.EXAMPLE
	AddWordTable -CustomObject $PSCustomObjectArray -Columns FirstName,LastName,EmailAddress -Headers "First Name","Last Name","Email Address"

	This example adds a table to the MS Word document, but only using the specified
	PSCustomObject note properties: FirstName, LastName and EmailAddress. If other note
	properties are present in the array of PSCustomObjects they will be ignored. The
	display names for each specified column header has been overridden to display a
	custom header. Note: the order of the header names must match the specified columns.
#>

Function AddWordTable
{
	[CmdletBinding()]
	Param
	(
		# Array of Hashtable (including table headers)
		[Parameter(Mandatory=$True, ValueFromPipelineByPropertyName=$True, ParameterSetName='Hashtable', Position=0)]
		[ValidateNotNullOrEmpty()] [System.Collections.Hashtable[]] $Hashtable,
		# Array of PSCustomObjects
		[Parameter(Mandatory=$True, ValueFromPipelineByPropertyName=$True, ParameterSetName='CustomObject', Position=0)]
		[ValidateNotNullOrEmpty()] [PSCustomObject[]] $CustomObject,
		# Array of Hashtable key names or PSCustomObject property names to include, in display order.
		# If not supplied then all Hashtable keys or all PSCustomObject properties will be displayed.
		[Parameter(ValueFromPipelineByPropertyName=$True)] [AllowNull()] [string[]] $Columns = $Null,
		# Array of custom table header strings in display order.
		[Parameter(ValueFromPipelineByPropertyName=$True)] [AllowNull()] [string[]] $Headers = $Null,
		# AutoFit table behavior.
		[Parameter(ValueFromPipelineByPropertyName=$True)] [AllowNull()] [int] $AutoFit = -1,
		# List view (no headers)
		[Switch] $List,
		# Grid lines
		[Switch] $NoGridLines,
		[Switch] $NoInternalGridLines,
		# Built-in Word table formatting style constant
		# Would recommend only $wdTableFormatContempory for normal usage (possibly $wdTableFormatList5 for List view)
		[Parameter(ValueFromPipelineByPropertyName=$True)] [int] $Format = 0
	)

	Begin 
	{
		Write-Debug ("Using parameter set '{0}'" -f $PSCmdlet.ParameterSetName);
		## Check if -Columns wasn't specified but -Headers were (saves some additional parameter sets!)
		If(($Null -eq $Columns) -and ($Null -ne $Headers)) 
		{
			Write-Warning "No columns specified and therefore, specified headers will be ignored.";
			$Columns = $Null;
		}
		ElseIf(($Null -ne $Columns) -and ($Null -ne $Headers)) 
		{
			## Check if number of specified -Columns matches number of specified -Headers
			If($Columns.Length -ne $Headers.Length) 
			{
				Write-Error "The specified number of columns does not match the specified number of headers.";
			}
		} ## end ElseIf
	} ## end Begin

	Process
	{
		## Build the Word table data string to be converted to a range and then a table later.
		[System.Text.StringBuilder] $WordRangeString = New-Object System.Text.StringBuilder;

		Switch ($PSCmdlet.ParameterSetName) 
		{
			'CustomObject' 
			{
				If($Null -eq $Columns) 
				{
					## Build the available columns from all available PSCustomObject note properties
					[string[]] $Columns = @();
					## Add each NoteProperty name to the array
					ForEach($Property in ($CustomObject | Get-Member -MemberType NoteProperty)) 
					{ 
						$Columns += $Property.Name; 
					}
				}

				## Add the table headers from -Headers or -Columns (except when in -List(view)
				If(-not $List) 
				{
					Write-Debug ("$(Get-Date): `t`tBuilding table headers");
					If($Null -ne $Headers) 
					{
                        [ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Headers));
					}
					Else 
					{ 
                        [ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Columns));
					}
				}

				## Iterate through each PSCustomObject
				Write-Debug ("$(Get-Date): `t`tBuilding table rows");
				ForEach($Object in $CustomObject) 
				{
					$OrderedValues = @();
					## Add each row item in the specified order
					ForEach($Column in $Columns) 
					{ 
						$OrderedValues += $Object.$Column; 
					}
					## Use the ordered list to add each column in specified order
					[ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $OrderedValues));
				} ## end ForEach
				Write-Debug ("$(Get-Date): `t`t`tAdded '{0}' table rows" -f ($CustomObject.Count));
			} ## end CustomObject

			Default 
			{   ## Hashtable
				If($Null -eq $Columns) 
				{
					## Build the available columns from all available hashtable keys. Hopefully
					## all Hashtables have the same keys (they should for a table).
					$Columns = $Hashtable[0].Keys;
				}

				## Add the table headers from -Headers or -Columns (except when in -List(view)
				If(-not $List) 
				{
					Write-Debug ("$(Get-Date): `t`tBuilding table headers");
					If($Null -ne $Headers) 
					{ 
						[ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Headers));
					}
					Else 
					{
						[ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Columns));
					}
				}
                
				## Iterate through each Hashtable
				Write-Debug ("$(Get-Date): `t`tBuilding table rows");
				ForEach($Hash in $Hashtable) 
				{
					$OrderedValues = @();
					## Add each row item in the specified order
					ForEach($Column in $Columns) 
					{ 
						$OrderedValues += $Hash.$Column; 
					}
					## Use the ordered list to add each column in specified order
					[ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $OrderedValues));
				} ## end ForEach

				Write-Debug ("$(Get-Date): `t`t`tAdded '{0}' table rows" -f $Hashtable.Count);
			} ## end default
		} ## end Switch

		## Create a MS Word range and set its text to our tab-delimited, concatenated string
		Write-Debug ("$(Get-Date): `t`tBuilding table range");
		$WordRange = $Script:Doc.Application.Selection.Range;
		$WordRange.Text = $WordRangeString.ToString();

		## Create hash table of named arguments to pass to the ConvertToTable method
		$ConvertToTableArguments = @{ Separator = [Microsoft.Office.Interop.Word.WdTableFieldSeparator]::wdSeparateByTabs; }

		## Negative built-in styles are not supported by the ConvertToTable method
		If($Format -ge 0) 
		{
			$ConvertToTableArguments.Add("Format", $Format);
			$ConvertToTableArguments.Add("ApplyBorders", $True);
			$ConvertToTableArguments.Add("ApplyShading", $True);
			$ConvertToTableArguments.Add("ApplyFont", $True);
			$ConvertToTableArguments.Add("ApplyColor", $True);
			If(!$List) 
			{ 
				$ConvertToTableArguments.Add("ApplyHeadingRows", $True); 
			}
			$ConvertToTableArguments.Add("ApplyLastRow", $True);
			$ConvertToTableArguments.Add("ApplyFirstColumn", $True);
			$ConvertToTableArguments.Add("ApplyLastColumn", $True);
		}

		## Invoke ConvertToTable method - with named arguments - to convert Word range to a table
		## See http://msdn.microsoft.com/en-us/library/office/aa171893(v=office.11).aspx
		Write-Debug ("$(Get-Date): `t`tConverting range to table");
		## Store the table reference just in case we need to set alternate row coloring
		$WordTable = $WordRange.GetType().InvokeMember(
			"ConvertToTable",                               # Method name
			[System.Reflection.BindingFlags]::InvokeMethod, # Flags
			$Null,                                          # Binder
			$WordRange,                                     # Target (self!)
			([Object[]]($ConvertToTableArguments.Values)),  ## Named argument values
			$Null,                                          # Modifiers
			$Null,                                          # Culture
			([String[]]($ConvertToTableArguments.Keys))     ## Named argument names
		);

		## Implement grid lines (will wipe out any existing formatting
		If($Format -lt 0) 
		{
			Write-Debug ("$(Get-Date): `t`tSetting table format");
			$WordTable.Style = $Format;
		}

		## Set the table autofit behavior
		If($AutoFit -ne -1) 
		{ 
			$WordTable.AutoFitBehavior($AutoFit); 
		}

		If(!$List)
		{
			#the next line causes the heading row to flow across page breaks
			$WordTable.Rows.First.Headingformat = $wdHeadingFormatTrue;
		}

		If(!$NoGridLines) 
		{
			$WordTable.Borders.InsideLineStyle = $wdLineStyleSingle;
			$WordTable.Borders.OutsideLineStyle = $wdLineStyleSingle;
		}
		If($NoGridLines) 
		{
			$WordTable.Borders.InsideLineStyle = $wdLineStyleNone;
			$WordTable.Borders.OutsideLineStyle = $wdLineStyleNone;
		}
		If($NoInternalGridLines) 
		{
			$WordTable.Borders.InsideLineStyle = $wdLineStyleNone;
			$WordTable.Borders.OutsideLineStyle = $wdLineStyleSingle;
		}

		Return $WordTable;

	} ## end Process
}

<#
.Synopsis
	Sets the format of one or more Word table cells
.DESCRIPTION
	This Function sets the format of one or more table cells, either from a collection
	of Word COM object cell references, an individual Word COM object cell reference or
	a hashtable containing Row and Column information.

	The font name, font size, bold, italic , underline and shading values can be used.
.EXAMPLE
	SetWordCellFormat -Hashtable $Coordinates -Table $TableReference -Bold

	This example sets all text to bold that is contained within the $TableReference
	Word table, using an array of hashtables. Each hashtable contain a pair of co-
	ordinates that is used to select the required cells. Note: the hashtable must
	contain the .Row and .Column key names. For example:
	@ { Row = 7; Column = 3 } to set the cell at row 7 and column 3 to bold.
.EXAMPLE
	$RowCollection = $Table.Rows.First.Cells
	SetWordCellFormat -Collection $RowCollection -Bold -Size 10

	This example sets all text to size 8 and bold for all cells that are contained
	within the first row of the table.
	Note: the $Table.Rows.First.Cells Returns a collection of Word COM cells objects
	that are in the first table row.
.EXAMPLE
	$ColumnCollection = $Table.Columns.Item(2).Cells
	SetWordCellFormat -Collection $ColumnCollection -BackgroundColor 255

	This example sets the background (shading) of all cells in the table's second
	column to red.
	Note: the $Table.Columns.Item(2).Cells Returns a collection of Word COM cells objects
	that are in the table's second column.
.EXAMPLE
	SetWordCellFormat -Cell $Table.Cell(17,3) -Font "Tahoma" -Color 16711680

	This example sets the font to Tahoma and the text color to blue for the cell located
	in the table's 17th row and 3rd column.
	Note: the $Table.Cell(17,3) Returns a single Word COM cells object.
#>

Function SetWordCellFormat 
{
	[CmdletBinding(DefaultParameterSetName='Collection')]
	Param (
		# Word COM object cell collection reference
		[Parameter(Mandatory=$true, ValueFromPipeline=$true, ParameterSetName='Collection', Position=0)] [ValidateNotNullOrEmpty()] $Collection,
		# Word COM object individual cell reference
		[Parameter(Mandatory=$true, ParameterSetName='Cell', Position=0)] [ValidateNotNullOrEmpty()] $Cell,
		# Hashtable of cell co-ordinates
		[Parameter(Mandatory=$true, ParameterSetName='Hashtable', Position=0)] [ValidateNotNullOrEmpty()] [System.Collections.Hashtable[]] $Coordinates,
		# Word COM object table reference
		[Parameter(Mandatory=$true, ParameterSetName='Hashtable', Position=1)] [ValidateNotNullOrEmpty()] $Table,
		# Font name
		[Parameter()] [AllowNull()] [string] $Font = $Null,
		# Font color
		[Parameter()] [AllowNull()] $Color = $Null,
		# Font size
		[Parameter()] [ValidateNotNullOrEmpty()] [int] $Size = 0,
		# Cell background color
		[Parameter()] [AllowNull()] [int]$BackgroundColor = $Null,
		# Force solid background color
		[Switch] $Solid,
		[Switch] $Bold,
		[Switch] $Italic,
		[Switch] $Underline
	)

	Begin 
	{
		Write-Debug ("Using parameter set '{0}'." -f $PSCmdlet.ParameterSetName);
	}

	Process 
	{
		Switch ($PSCmdlet.ParameterSetName) 
		{
			'Collection' {
				ForEach($Cell in $Collection) 
				{
					If($Null -ne $BackgroundColor) { $Cell.Shading.BackgroundPatternColor = $BackgroundColor; }
					If($Bold) { $Cell.Range.Font.Bold = $true; }
					If($Italic) { $Cell.Range.Font.Italic = $true; }
					If($Underline) { $Cell.Range.Font.Underline = 1; }
					If($Null -ne $Font) { $Cell.Range.Font.Name = $Font; }
					If($Null -ne $Color) { $Cell.Range.Font.Color = $Color; }
					If($Size -ne 0) { $Cell.Range.Font.Size = $Size; }
					If($Solid) { $Cell.Shading.Texture = 0; } ## wdTextureNone
				} # end ForEach
			} # end Collection
			'Cell' 
			{
				If($Bold) { $Cell.Range.Font.Bold = $true; }
				If($Italic) { $Cell.Range.Font.Italic = $true; }
				If($Underline) { $Cell.Range.Font.Underline = 1; }
				If($Null -ne $Font) { $Cell.Range.Font.Name = $Font; }
				If($Null -ne $Color) { $Cell.Range.Font.Color = $Color; }
				If($Size -ne 0) { $Cell.Range.Font.Size = $Size; }
				If($Null -ne $BackgroundColor) { $Cell.Shading.BackgroundPatternColor = $BackgroundColor; }
				If($Solid) { $Cell.Shading.Texture = 0; } ## wdTextureNone
			} # end Cell
			'Hashtable' 
			{
				ForEach($Coordinate in $Coordinates) 
				{
					$Cell = $Table.Cell($Coordinate.Row, $Coordinate.Column);
					If($Bold) { $Cell.Range.Font.Bold = $true; }
					If($Italic) { $Cell.Range.Font.Italic = $true; }
					If($Underline) { $Cell.Range.Font.Underline = 1; }
					If($Null -ne $Font) { $Cell.Range.Font.Name = $Font; }
					If($Null -ne $Color) { $Cell.Range.Font.Color = $Color; }
					If($Size -ne 0) { $Cell.Range.Font.Size = $Size; }
					If($Null -ne $BackgroundColor) { $Cell.Shading.BackgroundPatternColor = $BackgroundColor; }
					If($Solid) { $Cell.Shading.Texture = 0; } ## wdTextureNone
				}
			} # end Hashtable
		} # end Switch
	} # end process
}

<#
.Synopsis
	Sets alternate row colors in a Word table
.DESCRIPTION
	This Function sets the format of alternate rows within a Word table using the
	specified $BackgroundColor. This Function is expensive (in performance terms) as
	it recursively sets the format on alternate rows. It would be better to pick one
	of the predefined table formats (if one exists)? Obviously the more rows, the
	longer it takes :'(

	Note: this Function is called by the AddWordTable Function if an alternate row
	format is specified.
.EXAMPLE
	SetWordTableAlternateRowColor -Table $TableReference -BackgroundColor 255

	This example sets every-other table (starting with the first) row and sets the
	background color to red (wdColorRed).
.EXAMPLE
	SetWordTableAlternateRowColor -Table $TableReference -BackgroundColor 39423 -Seed Second

	This example sets every other table (starting with the second) row and sets the
	background color to light orange (weColorLightOrange).
#>

Function SetWordTableAlternateRowColor 
{
	[CmdletBinding()]
	Param (
		# Word COM object table reference
		[Parameter(Mandatory=$true, ValueFromPipeline=$true, Position=0)] [ValidateNotNullOrEmpty()] $Table,
		# Alternate row background color
		[Parameter(Mandatory=$true, Position=1)] [ValidateNotNull()] [int] $BackgroundColor,
		# Alternate row starting seed
		[Parameter(ValueFromPipelineByPropertyName=$true, Position=2)] [ValidateSet('First','Second')] [string] $Seed = 'First'
	)

	Process 
	{
		$StartDateTime = Get-Date;
		Write-Debug ("{0}: `t`tSetting alternate table row colors.." -f $StartDateTime);

		## Determine the row seed (only really need to check for 'Second' and default to 'First' otherwise
		If($Seed.ToLower() -eq 'second') 
		{ 
			$StartRowIndex = 2; 
		}
		Else 
		{ 
			$StartRowIndex = 1; 
		}

		For($AlternateRowIndex = $StartRowIndex; $AlternateRowIndex -lt $Table.Rows.Count; $AlternateRowIndex += 2) 
		{ 
			$Table.Rows.Item($AlternateRowIndex).Shading.BackgroundPatternColor = $BackgroundColor;
		}

		## I've put verbose calls in here we can see how expensive this Functionality actually is.
		$EndDateTime = Get-Date;
		$ExecutionTime = New-TimeSpan -Start $StartDateTime -End $EndDateTime;
		Write-Debug ("{0}: `t`tDone setting alternate row style color in '{1}' seconds" -f $EndDateTime, $ExecutionTime.TotalSeconds);
	}
}
#endregion

#region PoSH prereqs
Function CheckOnPoSHPrereqs
{
	Write-Verbose "$(Get-Date -Format G): Checking for McliPSSnapin"
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
			Write-Verbose "$(Get-Date -Format G): Searching for the 32-bit .Net V2 snapin"
			$installutil = $env:systemroot + '\Microsoft.NET\Framework\v2.0.50727\installutil.exe'
			If(Test-Path $installutil -EA 0)
			{
				Write-Verbose "$(Get-Date -Format G): `tAttempting to register the 32-bit .Net V2 snapin"
				&$installutil $PVSDLLPath > $Null
			
				If(!$?)
				{
					Write-Verbose "$(Get-Date -Format G): `t`tUnable to register the 32-bit V2 PowerShell Snap-in."
				}
				Else
				{
					Write-Verbose "$(Get-Date -Format G): `t`tRegistered the 32-bit V2 PowerShell Snap-in."
				}
			}
			Else
			{
				Write-Verbose "$(Get-Date -Format G): `tNo 32-bit .Net V2 snapin found"
			}
	
			Write-Verbose "$(Get-Date -Format G): Searching for the 64-bit .Net V2 snapin"
			$installutil = $env:systemroot + '\Microsoft.NET\Framework64\v2.0.50727\installutil.exe'
			If(Test-Path $installutil -EA 0)
			{
				Write-Verbose "$(Get-Date -Format G): `tAttempting to register the 64-bit .Net V2 snapin"
				&$installutil $PVSDLLPath > $Null
			
				If(!$?)
				{
					Write-Verbose "$(Get-Date -Format G): `t`tUnable to register the 64-bit V2 PowerShell Snap-in."
				}
				Else
				{
					Write-Verbose "$(Get-Date -Format G): `t`tRegistered the 64-bit V2 PowerShell Snap-in."
				}
			}
			Else
			{
				Write-Verbose "$(Get-Date -Format G): `tNo 64-bit .Net V2 snapin found"
			}
		}
		Else
		{
			Write-Verbose "$(Get-Date -Format G): Unable to find $PVSDLLPath"
		}
		
		If(Test-Path $PVSDLLPath -EA 0)
		{
			Write-Verbose "$(Get-Date -Format G): Searching for the 32-bit .Net V4 snapin"
			$installutil = $env:systemroot + '\Microsoft.NET\Framework\v4.0.30319\installutil.exe'
			If(Test-Path $installutil -EA 0)
			{
				Write-Verbose "$(Get-Date -Format G): `tAttempting to register the 32-bit .Net V4 snapin"
				&$installutil $PVSDLLPath > $Null
			
				If(!$?)
				{
					Write-Verbose "$(Get-Date -Format G): `t`tUnable to register the 32-bit V4 PowerShell Snap-in."
				}
				Else
				{
					Write-Verbose "$(Get-Date -Format G): `t`tRegistered the 32-bit V4 PowerShell Snap-in."
				}
			}
			Else
			{
				Write-Verbose "$(Get-Date -Format G): `tNo 32-bit .Net V4 snapin found"
			}
	
			Write-Verbose "$(Get-Date -Format G): Searching for the 64-bit .Net V4 snapin"
			$installutil = $env:systemroot + '\Microsoft.NET\Framework64\v4.0.30319\installutil.exe'
			If(Test-Path $installutil -EA 0)
			{
				Write-Verbose "$(Get-Date -Format G): `tAttempting to register the 64-bit .Net V4 snapin"
				&$installutil $PVSDLLPath > $Null
			
				If(!$?)
				{
					Write-Verbose "$(Get-Date -Format G): `t`tUnable to register the 64-bit V4 PowerShell Snap-in."
				}
				Else
				{
					Write-Verbose "$(Get-Date -Format G): `t`tRegistered the 64-bit V4 PowerShell Snap-in."
				}
			}
			Else
			{
				Write-Verbose "$(Get-Date -Format G): `tNo 64-bit .Net V4 snapin found"
			}
		}
		Else
		{
			Write-Verbose "$(Get-Date -Format G): Unable to find $PVSDLLPath"
		}
	
		Write-Verbose "$(Get-Date -Format G): Rechecking for McliPSSnapin"
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
			Write-Verbose "$(Get-Date -Format G): Citrix PowerShell Snap-ins detected at $PVSDLLPath"
		}
	}
	Else
	{
		Write-Verbose "$(Get-Date -Format G): Citrix PowerShell Snap-ins detected."
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
				Write-Verbose "$(Get-Date -Format G): This script is being run remotely against server $($AdminAddress)"
				If(![System.String]::IsNullOrEmpty($User))
				{
					Write-Verbose "$(Get-Date -Format G): User=$($User)"
					Write-Verbose "$(Get-Date -Format G): Domain=$($Domain)"
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
	
	Write-Verbose "$(Get-Date -Format G): Verifying PVS SOAP and Stream Services are running on $tmp"

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
	Write-Verbose "$(Get-Date -Format G): Getting PVS version info"

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
	Write-Verbose "$(Get-Date -Format G): Build PVS farm values"
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
}
#endregion

#region show script options
Function ShowScriptOptions
{
	Write-Verbose "$(Get-Date -Format G): "
	Write-Verbose "$(Get-Date -Format G): "
	Write-Verbose "$(Get-Date -Format G): AddDateTime         : $AddDateTime"
	Write-Verbose "$(Get-Date -Format G): AdminAddress        : $AdminAddress"
	If($MSWORD -or $PDF)
	{
		Write-Verbose "$(Get-Date -Format G): Company Name        : $Script:CoName"
		Write-Verbose "$(Get-Date -Format G): Company Address     : $CompanyAddress"
		Write-Verbose "$(Get-Date -Format G): Company Email       : $CompanyEmail"
		Write-Verbose "$(Get-Date -Format G): Company Fax         : $CompanyFax"
		Write-Verbose "$(Get-Date -Format G): Company Phone       : $CompanyPhone"
		Write-Verbose "$(Get-Date -Format G): Cover Page          : $CoverPage"
	}
	Write-Verbose "$(Get-Date -Format G): CSV                 : $CSV"
	Write-Verbose "$(Get-Date -Format G): Dev                 : $Dev"
	If($Dev)
	{
		Write-Verbose "$(Get-Date -Format G): DevErrorFile        : $Script:DevErrorFile"
	}
	Write-Verbose "$(Get-Date -Format G): Domain              : $Domain"
	If($HTML)
	{
		Write-Verbose "$(Get-Date -Format G): HTMLFilename        : $Script:HTMLFilename"
	}
	If($MSWord)
	{
		Write-Verbose "$(Get-Date -Format G): WordFilename        : $Script:WordFilename"
	}
	If($PDF)
	{
		Write-Verbose "$(Get-Date -Format G): PDFFilename         : $Script:PDFFilename"
	}
	If($Text)
	{
		Write-Verbose "$(Get-Date -Format G): TextFilename        : $Script:TextFilename"
	}
	Write-Verbose "$(Get-Date -Format G): Folder              : $Folder"
	Write-Verbose "$(Get-Date -Format G): From                : $From"
	Write-Verbose "$(Get-Date -Format G): Log                 : $Log"
	Write-Verbose "$(Get-Date -Format G): PVS Version         : $Script:PVSFullVersion"
	Write-Verbose "$(Get-Date -Format G): Report Footer       : $ReportFooter"
	Write-Verbose "$(Get-Date -Format G): Save As HTML        : $HTML"
	Write-Verbose "$(Get-Date -Format G): Save As PDF         : $PDF"
	Write-Verbose "$(Get-Date -Format G): Save As TEXT        : $TEXT"
	Write-Verbose "$(Get-Date -Format G): Save As WORD        : $MSWORD"
	Write-Verbose "$(Get-Date -Format G): ScriptInfo          : $ScriptInfo"
	Write-Verbose "$(Get-Date -Format G): Smtp Port           : $SmtpPort"
	Write-Verbose "$(Get-Date -Format G): Smtp Server         : $SmtpServer"
	Write-Verbose "$(Get-Date -Format G): Title               : $Script:Title"
	Write-Verbose "$(Get-Date -Format G): To                  : $To"
	Write-Verbose "$(Get-Date -Format G): Use SSL             : $UseSSL"
	Write-Verbose "$(Get-Date -Format G): User                : $User"
	If($MSWORD -or $PDF)
	{
		Write-Verbose "$(Get-Date -Format G): Username            : $UserName"
	}
	Write-Verbose "$(Get-Date -Format G): "
	Write-Verbose "$(Get-Date -Format G): OS Detected         : $Script:RunningOS"
	Write-Verbose "$(Get-Date -Format G): PoSH version        : $($Host.Version)"
	Write-Verbose "$(Get-Date -Format G): PSCulture           : $PSCulture"
	Write-Verbose "$(Get-Date -Format G): PSUICulture         : $PSUICulture"
	If($MSWORD -or $PDF)
	{
		Write-Verbose "$(Get-Date -Format G): Word language       : $Script:WordLanguageValue"
		Write-Verbose "$(Get-Date -Format G): Word version        : $Script:WordProduct"
	}
	Write-Verbose "$(Get-Date -Format G): "
	Write-Verbose "$(Get-Date -Format G): Script start        : $Script:StartTime"
	Write-Verbose "$(Get-Date -Format G): "
	Write-Verbose "$(Get-Date -Format G): "
}
#endregion

Function validStateProp( [object] $object, [string] $topLevel, [string] $secondLevel )
{
	#function created 8-jan-2015 by Michael B. Smith
	if( $object )
	{
		If( ( Get-Member -Name $topLevel -InputObject $object ) )
		{
			If( ( Get-Member -Name $secondLevel -InputObject $object.$topLevel ) )
			{
				Return $True
			}
		}
	}
	Return $False
}

Function validObject( [object] $object, [string] $topLevel )
{
	#function created 8-jan-2014 by Michael B. Smith
	If( $object )
	{
		If((Get-Member -Name $topLevel -InputObject $object))
		{
			Return $True
		}
	}
	Return $False
}

Function SaveandCloseDocumentandShutdownWord
{
	#bug fix 1-Apr-2014
	#reset Grammar and Spelling options back to their original settings
	$Script:Word.Options.CheckGrammarAsYouType = $Script:CurrentGrammarOption
	$Script:Word.Options.CheckSpellingAsYouType = $Script:CurrentSpellingOption

	Write-Verbose "$(Get-Date -Format G): Save and Close document and Shutdown Word"
	If($Script:WordVersion -eq $wdWord2010)
	{
		#the $saveFormat below passes StrictMode 2
		#I found this at the following two links
		#http://msdn.microsoft.com/en-us/library/microsoft.office.interop.word.wdsaveformat(v=office.14).aspx
		If($PDF)
		{
			Write-Verbose "$(Get-Date -Format G): Saving as DOCX file first before saving to PDF"
		}
		Else
		{
			Write-Verbose "$(Get-Date -Format G): Saving DOCX file"
		}
		Write-Verbose "$(Get-Date -Format G): Running $($Script:WordProduct) and detected operating system $($Script:RunningOS)"
		$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatDocumentDefault")
		$Script:Doc.SaveAs([REF]$Script:WordFileName, [ref]$SaveFormat)
		If($PDF)
		{
			Write-Verbose "$(Get-Date -Format G): Now saving as PDF"
			$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatPDF")
			$Script:Doc.SaveAs([REF]$Script:PDFFileName, [ref]$saveFormat)
		}
	}
	ElseIf($Script:WordVersion -eq $wdWord2013 -or $Script:WordVersion -eq $wdWord2016)
	{
		If($PDF)
		{
			Write-Verbose "$(Get-Date -Format G): Saving as DOCX file first before saving to PDF"
		}
		Else
		{
			Write-Verbose "$(Get-Date -Format G): Saving DOCX file"
		}
		Write-Verbose "$(Get-Date -Format G): Running $($Script:WordProduct) and detected operating system $($Script:RunningOS)"
		$Script:Doc.SaveAs2([REF]$Script:WordFileName, [ref]$wdFormatDocumentDefault)
		If($PDF)
		{
			Write-Verbose "$(Get-Date -Format G): Now saving as PDF"
			$Script:Doc.SaveAs([REF]$Script:PDFFileName, [ref]$wdFormatPDF)
		}
	}

	Write-Verbose "$(Get-Date -Format G): Closing Word"
	$Script:Doc.Close()
	$Script:Word.Quit()
	Write-Verbose "$(Get-Date -Format G): System Cleanup"
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Script:Word) | Out-Null
	If(Test-Path variable:global:word)
	{
		Remove-Variable -Name word -Scope Global 4>$Null
	}
	$SaveFormat = $Null
	[gc]::collect() 
	[gc]::WaitForPendingFinalizers()
	
	#is the winword Process still running? kill it

	#find out our session (usually "1" except on TS/RDC or Citrix)
	$SessionID = (Get-Process -PID $PID).SessionId

	#Find out if winword running in our session
	$wordprocess = ((Get-Process 'WinWord' -ea 0) | Where-Object {$_.SessionId -eq $SessionID}) | Select-Object -Property Id 
	If( $wordprocess -and $wordprocess.Id -gt 0)
	{
		Write-Verbose "$(Get-Date -Format G): WinWord Process is still running. Attempting to stop WinWord Process # $($wordprocess.Id)"
		Stop-Process $wordprocess.Id -EA 0
	}
}

Function SetupText
{
	Write-Verbose "$(Get-Date -Format G): Setting up Text"

	[System.Text.StringBuilder] $global:Output = New-Object System.Text.StringBuilder( 16384 )

	If(!$AddDateTime)
	{
		[string]$Script:TextFileName = "$($Script:pwdpath)\$($OutputFileName).txt"
	}
	ElseIf($AddDateTime)
	{
		[string]$Script:TextFileName = "$($Script:pwdpath)\$($OutputFileName)_$(Get-Date -f yyyy-MM-dd_HHmm).txt"
	}
}

Function SaveandCloseTextDocument
{
	Write-Verbose "$(Get-Date -Format G): Saving Text file"
	Line 0 ""
	Line 0 "Report Complete"
	Write-Output $global:Output.ToString() | Out-File $Script:TextFileName 4>$Null
}

Function SaveandCloseHTMLDocument
{
	Write-Verbose "$(Get-Date -Format G): Saving HTML file"
	WriteHTMLLine 0 0 ""
	WriteHTMLLine 0 0 "Report Complete"
	Out-File -FilePath $Script:HTMLFileName -Append -InputObject "<p></p></body></html>" 4>$Null
}

Function SetFilenames
{
	Param([string]$OutputFileName)
	
	If($MSWord -or $PDF)
	{
		CheckWordPreReq
		
		SetupWord
	}
	If($Text)
	{
		SetupText
	}
	If($HTML)
	{
		SetupHTML
	}
	ShowScriptOptions
}

Function OutputReportFooter
{
	<#
	Report Footer
		Report information:
			Created with: <Script Name> - Release Date: <Script Release Date>
			Script version: <Script Version>
			Started on <Date Time in Local Format>
			Elapsed time: nn days, nn hours, nn minutes, nn.nn seconds
			Ran from domain <Domain Name> by user <Username>
			Ran from the folder <Folder Name>

	Script Name and Script Release date are script-specific variables.
	Script version is a script variable.
	Start Date Time in Local Format is a script variable.
	Domain Name is $env:USERDNSDOMAIN.
	Username is $env:USERNAME.
	Folder Name is a script variable.
	#>

	$runtime = $(Get-Date) - $Script:StartTime
	$Str = [string]::format("{0} days, {1} hours, {2} minutes, {3}.{4} seconds",
		$runtime.Days,
		$runtime.Hours,
		$runtime.Minutes,
		$runtime.Seconds,
		$runtime.Milliseconds)

	If($MSWORD -or $PDF)
	{
		$Script:selection.InsertNewPage()
		WriteWordLine 1 0 "Report Footer"
		WriteWordLine 2 0 "Report Information:"
		WriteWordLine 0 1 "Created with: $Script:ScriptName - Release Date: $Script:ReleaseDate"
		WriteWordLine 0 1 "Script version: $Script:MyVersion"
		WriteWordLine 0 1 "Started on $Script:StartTime"
		WriteWordLine 0 1 "Elapsed time: $Str"
		WriteWordLine 0 1 "Ran from domain $env:USERDNSDOMAIN by user $env:USERNAME"
		WriteWordLine 0 1 "Ran from the folder $Script:pwdpath"
	}
	If($Text)
	{
		Line 0 "///  Report Footer  \\\"
		Line 1 "Report Information:"
		Line 2 "Created with: $Script:ScriptName - Release Date: $Script:ReleaseDate"
		Line 2 "Script version: $Script:MyVersion"
		Line 2 "Started on $Script:StartTime"
		Line 2 "Elapsed time: $Str"
		Line 2 "Ran from domain $env:USERDNSDOMAIN by user $env:USERNAME"
		Line 2 "Ran from the folder $Script:pwdpath"
	}
	If($HTML)
	{
		WriteHTMLLine 1 0 "///&nbsp;&nbsp;Report Footer&nbsp;&nbsp;\\\"
		WriteHTMLLine 2 0 "Report Information:"
		WriteHTMLLine 0 1 "Created with: $Script:ScriptName - Release Date: $Script:ReleaseDate"
		WriteHTMLLine 0 1 "Script version: $Script:MyVersion"
		WriteHTMLLine 0 1 "Started on $Script:StartTime"
		WriteHTMLLine 0 1 "Elapsed time: $Str"
		WriteHTMLLine 0 1 "Ran from domain $env:USERDNSDOMAIN by user $env:USERNAME"
		WriteHTMLLine 0 1 "Ran from the folder $Script:pwdpath"
	}
}

Function ProcessDocumentOutput
{
	If($MSWORD -or $PDF)
	{
		SaveandCloseDocumentandShutdownWord
	}
	If($Text)
	{
		SaveandCloseTextDocument
	}
	If($HTML)
	{
		SaveandCloseHTMLDocument
	}

	$GotFile = $False

	If($MSWord)
	{
		If(Test-Path "$($Script:WordFileName)")
		{
			Write-Verbose "$(Get-Date -Format G): $($Script:WordFileName) is ready for use"
			$GotFile = $True
		}
		Else
		{
			Write-Error "Unable to save the output file, $($Script:WordFileName)"
		}
	}
	If($PDF)
	{
		If(Test-Path "$($Script:PDFFileName)")
		{
			Write-Verbose "$(Get-Date -Format G): $($Script:PDFFileName) is ready for use"
			$GotFile = $True
		}
		Else
		{
			Write-Error "Unable to save the output file, $($Script:PDFFileName)"
		}
	}
	If($Text)
	{
		If(Test-Path "$($Script:TextFileName)")
		{
			Write-Verbose "$(Get-Date -Format G): $($Script:TextFileName) is ready for use"
			$GotFile = $True
		}
		Else
		{
			Write-Error "Unable to save the output file, $($Script:TextFileName)"
		}
	}
	If($HTML)
	{
		If(Test-Path "$($Script:HTMLFileName)")
		{
			Write-Verbose "$(Get-Date -Format G): $($Script:HTMLFileName) is ready for use"
			$GotFile = $True
		}
		Else
		{
			Write-Error "Unable to save the output file, $($Script:HTMLFileName)"
		}
	}
	
	#email output file if requested
	If($GotFile -and ![System.String]::IsNullOrEmpty( $SmtpServer ))
	{
		$emailattachments = @()
		If($MSWord)
		{
			$emailAttachments += $Script:WordFileName
		}
		If($PDF)
		{
			$emailAttachments += $Script:PDFFileName
		}
		If($Text)
		{
			$emailAttachments += $Script:TextFileName
		}
		If($HTML)
		{
			$emailAttachments += $Script:HTMLFileName
		}
		SendEmail $emailAttachments
	}
}

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

Function OutputauthGroups
{
	Param([object] $authGroups)
	
	If($MSWord -or $PDF)
	{
		[System.Collections.Hashtable[]] $AuthWordTable = @();
	}
	If($HTML)
	{
		$rowdata = @()
	}

	ForEach($Group in $authgroups)
	{
		If($Group.authGroupName)
		{
			If($MSword -or $PDF)
			{
				$WordTableRowHash = @{Name = $Group.authGroupName;}
				$AuthWordTable += $WordTableRowHash;
			}
			If($Text)
			{
				Line 2 $Group.authGroupName
			}
			If($HTML)
			{
				$rowdata += @(,(
				$Group.authGroupName,$htmlwhite))
			}
		}
	}
	
	If($MSword -or $PDF)
	{
		If($AuthWordTable.Count -gt 0)
		{
			$Table = AddWordTable -Hashtable $AuthWordTable `
			-Columns Name `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitContent;

			SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""
		}
	}
	If($Text)
	{
		Line 0 ""
	}
	If($HTML)
	{
		$columnHeaders = @(
		'Name',($global:htmlsb))
		
		$msg = ""
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
	}
}

Function OutputWarning
{
	Param([string] $txt)
	Write-Warning $txt
	If($MSWord -or $PDF)
	{
		WriteWordLine 0 1 $txt
		WriteWordLine 0 0 ""
	}
	If($Text)
	{
		Line 1 $txt
		Line 0 ""
	}
	If($HTML)
	{
		WriteHTMLLine 0 1 $txt
		WriteHTMLLine 0 0 ""
	}
}

Function OutputNotice
{
	Param([string] $txt)
	If($MSWord -or $PDF)
	{
		WriteWordLine 0 0 $txt
		WriteWordLIne 0 0 ""
	}
	If($Text)
	{
		Line 0 $txt
		Line 0 ""
	}
	If($HTML)
	{
		WriteHTMLLine 0 0 $txt
	}
}

Function ProcessPVSFarm
{
	Write-Verbose "$(Get-Date -Format G): Processing PVS Farm Information"

	$LicenseServerIPAddress = Get-IPAddress $Script:farm.licenseServer #added in V1.16
	
	#V1.17 see if the database server names contain an instance name. If so, remove it
	#V1.18 add test for port number - bug found by Johan Parlevliet 
	#V1.18 see if the database server names contain a port number. If so, remove it
	#V1.18 optimized code supplied by MBS
	$dbServer = $Script:farm.databaseServerName
	If( ( $inx = $dbServer.IndexOfAny( ',\' ) ) -ge 0 )
	{
		#strip the instance name and/or port name, if present
		Write-Verbose "$(Get-Date -Format G): Removing '$( $dbServer.SubString( $inx ) )' from SQL server name to get IP address"
		$dbServer = $dbServer.SubString( 0, $inx )
		Write-Verbose "$(Get-Date -Format G): dbServer now '$dbServer'"
	}
	$SQLServerIPAddress = Get-IPAddress $dbServer #added in V1.16
	
	$dbServer = $Script:farm.failoverPartnerServerName
	If( ( $inx = $dbServer.IndexOfAny( ',\' ) ) -ge 0 )
	{
		#strip the instance name and/or port name, if present
		Write-Verbose "$(Get-Date -Format G): Removing '$( $dbServer.SubString( $inx ) )' from SQL server name to get IP address"
		$dbServer = $dbServer.SubString( 0, $inx )
		Write-Verbose "$(Get-Date -Format G): dbServer now '$dbServer'"
	}
	$FailoverSQLServerIPAddress = Get-IPAddress $dbServer #added in V1.16
	
	#general tab
	If($MSWord -or $PDF)
	{
		$selection.InsertNewPage()
		WriteWordLine 1 0 "PVS Farm Information"
		WriteWordLine 2 0 "General"
		[System.Collections.Hashtable[]] $ScriptInformation = @()
		$ScriptInformation += @{ Data = "PVS Farm Name"; Value = $Script:farm.farmName; }
		$ScriptInformation += @{ Data = "PVS Version"; Value = $Script:PVSFullVersion; }
		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitContent;

		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If($Text)
	{
		Line 0 "PVS Farm Information"
		Line 1 "General"
		Line 2 "PVS Farm Name`t: " $Script:farm.farmName
		Line 2 "Version`t`t: " $Script:PVSFullVersion
		Line 0 ""
	}
	If($HTML)
	{
		WriteHTMLLine 1 0 "PVS Farm Information"
		WriteHTMLLine 2 0 "General"
		$rowdata = @()
		$columnHeaders = @("PVS Farm Name",($global:htmlsb),$Script:farm.farmName,$htmlwhite)
		$rowdata += @(,('PVS Version',($global:htmlsb),$Script:PVSFullVersion,$htmlwhite))
		
		$msg = ""
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
	}
	
	#security tab
	Write-Verbose "$(Get-Date -Format G): `tProcessing Security Tab"
	If($MSWord -or $PDF)
	{
		WriteWordLine 2 0 "Security"
		WriteWordLine 0 0 "Groups with 'Farm Administrator' access"
	}
	If($Text)
	{
		Line 1 "Security"
		Line 2 "Groups with Farm Administrator access:"
	}
	If($HTML)
	{
		WriteHTMLLine 2 0 "Security"
		WriteHTMLLine 3 0 "Groups with Farm Administrator access:"
	}

	#build security tab values
	$GetWhat = "authgroup"
	$GetParam = "farm = 1"
	$ErrorTxt = "Groups with Farm Administrator access"
	$authgroups = BuildPVSObject $GetWhat $GetParam $ErrorTxt

	If($? -and $Null -ne $AuthGroups)
	{
		OutputauthGroups $authGroups
	}
	ElseIf($? -and $Null -eq $AuthGroups)
	{
		$txt = "There are no Farm authorization groups"
		OutputNotice $txt
	}
	Else
	{
		$txt = "Unable to retrieve Farm authorization groups"
		OutputWarning $txt
	}

	#groups tab
	Write-Verbose "$(Get-Date -Format G): `tProcessing Groups Tab"
	If($MSWord -or $PDF)
	{
		WriteWordLine 2 0 "Groups"
		WriteWordLine 0 0 "All the Security Groups that can be assigned access rights"
	}
	If($Text)
	{
		Line 1 "Groups"
		Line 2 "All the Security Groups that can be assigned access rights:"
	}
	If($HTML)
	{
		WriteHTMLLine 2 0 "Groups"
		WriteHTMLLine 3 0 "All the Security Groups that can be assigned access rights:"
	}

	$GetWhat = "authgroup"
	$GetParam = ""
	$ErrorTxt = "Security Groups information"
	$authgroups = BuildPVSObject $GetWhat $GetParam $ErrorTxt

	If($? -and $Null -ne $AuthGroups)
	{
		OutputauthGroups $authGroups
	}
	ElseIf($? -and $Null -eq $AuthGroups)
	{
		$txt = "There are no authorization groups"
		OutputNotice $txt
	}
	Else
	{
		$txt = "Unable to retrieve authorization groups"
		OutputWarning $txt
	}

	Write-Verbose "$(Get-Date -Format G): `tProcessing Licensing Tab"
	
	If($Script:farm.licenseTradeUp -eq "1" -or $Script:farm.licenseTradeUp -eq $True)
	{
		$DatacenterLicense = "Yes"
	}
	Else
	{
		$DatacenterLicense = "No"
	}
	
	If($MSWord -or $PDF)
	{
		WriteWordLine 2 0 "Licensing"
		[System.Collections.Hashtable[]] $ScriptInformation = @()
		$ScriptInformation += @{ Data = "License server name"; Value = $Script:farm.licenseServer; }
		$ScriptInformation += @{ Data = "License server IP"; Value = $LicenseServerIPAddress; }
		$ScriptInformation += @{ Data = "License server port"; Value = $Script:farm.licenseServerPort; }
		If($Script:PVSFullVersion -ge "7.19")
		{
			$ScriptInformation += @{ Data = "Citrix Provisioning license type"; Value = ""; }
			If($Script:farm.LicenseSKU -eq 0)
			{
				$ScriptInformation += @{ Data = "     On-Premises"; Value = "Yes"; }
				$ScriptInformation += @{ Data = "          Use Datacenter licenses for desktops if no Desktop licenses are available"; Value = $DatacenterLicense; }
				$ScriptInformation += @{ Data = "     Cloud"; Value = "No"; }
			}
			ElseIf($Script:farm.LicenseSKU -eq 1)
			{
				$ScriptInformation += @{ Data = "     On-Premises"; Value = "No"; }
				$ScriptInformation += @{ Data = "          Use Datacenter licenses for desktops if no Desktop licenses are available"; Value = $DatacenterLicense; }
				$ScriptInformation += @{ Data = "     Cloud"; Value = "Yes"; }
			}
			Else
			{
				$ScriptInformation += @{ Data = "     On-Premises"; Value = "ERROR: Unable to determine the PVS License SKU Tpe"; }
			}
		}
		ElseIf($Script:PVSFullVersion -ge "7.13")
		{
			$ScriptInformation += @{ Data = "Use Datacenter licenses for desktops if no Desktop licenses are available"; Value = $DatacenterLicense; }
		}
		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitContent;

		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If($Text)
	{
		Line 1 "Licensing"
		Line 2 "License server name`t: " $Script:farm.licenseServer
		Line 2 "License server IP`t: " $LicenseServerIPAddress
		Line 2 "License server port`t: " $Script:farm.licenseServerPort
		If($Script:PVSFullVersion -ge "7.19")
		{
			Line 2 "Citrix Provisioning license type" ""
			If($Script:farm.LicenseSKU -eq 0)
			{
				Line 3 "On-Premises`t: " "Yes"
				Line 4 "Use Datacenter licenses for desktops if no Desktop licenses are available: " $DatacenterLicense
				Line 3 "Cloud`t`t: " "No"
			}
			ElseIf($Scrpt:farm.LicenseSKU -eq 1)
			{
				Line 3 "On-Premises`t: " "No"
				Line 4 "Use Datacenter licenses for desktops if no Desktop licenses are available: " $DatacenterLicense
				Line 3 "Cloud`t`t: " "Yes"
			}
			Else
			{
				Line 3 "On-Premises`t: " "ERROR: Unable to determine the PVS License SKU Tpe"
			}
		}
		ElseIf($Script:PVSFullVersion -ge "7.13")
		{
			Line 2 "Use Datacenter licenses for desktops if no Desktop licenses are available: " $DatacenterLicense
		}
		Line 0 ""
	}
	If($HTML)
	{
		WriteHTMLLine 2 0 "Licensing"
		$rowdata = @()
		$columnHeaders = @("License server name",($global:htmlsb),$Script:farm.licenseServer,$htmlwhite)
		$rowdata += @(,('License server IP',($global:htmlsb),$LicenseServerIPAddress,$htmlwhite))
		$rowdata += @(,('License server port',($global:htmlsb),$Script:farm.licenseServerPort,$htmlwhite))
		If($Script:PVSFullVersion -ge "7.19")
		{
			$rowdata += @(,("Citrix Provisioning license type",($global:htmlsb),"",$htmlwhite))
			If($Script:farm.LicenseSKU -eq 0)
			{
				$rowdata += @(,("     On-Premises",($global:htmlsb),"Yes",$htmlwhite))
				$rowdata += @(,("          Use Datacenter licenses for desktops if no Desktop licenses are available",($global:htmlsb),$DatacenterLicense,$htmlwhite))
				$rowdata += @(,("     Cloud",($global:htmlsb),"No",$htmlwhite))
			}
			ElseIf($Script:farm.LicenseSKU -eq 1)
			{
				$rowdata += @(,("     On-Premises",($global:htmlsb),"No",$htmlwhite))
				$rowdata += @(,("          Use Datacenter licenses for desktops if no Desktop licenses are available",($global:htmlsb),$DatacenterLicense,$htmlwhite))
				$rowdata += @(,("     Cloud",($global:htmlsb),"Yes",$htmlwhite))
			}
			Else
			{
				$rowdata += @(,("     On-Premises",($global:htmlsb),"ERROR: Unable to determine the PVS License SKU Tpe",$htmlwhite))
			}
		}
		ElseIf($Script:PVSFullVersion -ge "7.13")
		{
			$rowdata += @(,('Use Datacenter licenses for desktops if no Desktop licenses are available',($global:htmlsb),$DatacenterLicense,$htmlwhite))
		}
		FormatHTMLTable "" "auto" -rowArray $rowdata -columnArray $columnHeaders
	}

	#options tab
	Write-Verbose "$(Get-Date -Format G): `tProcessing Options Tab"
	If($Script:farm.auditingEnabled -ne "1")
	{
		$obj1 = [PSCustomObject] @{
			ItemText = "Auditing is not enabled"
		}
		$null = $Script:ItemsToReview.Add($obj1)
	}
	If($Script:farm.offlineDatabaseSupportEnabled -ne "1")
	{
		$obj1 = [PSCustomObject] @{
			ItemText = "Offline database support is not enabled"
		}
		$null = $Script:ItemsToReview.Add($obj1)
	}
	
	If($Script:farm.autoAddEnabled -eq "1")
	{
		$Script:farmautoAddEnabled = $True
	}
	Else
	{
		$Script:farmautoAddEnabled = $False
	}	
	If($Script:farm.auditingEnabled -eq "1")
	{
		$Script:farmauditingEnabled = $True
	}
	Else
	{
		$Script:farmauditingEnabled = $False
	}
	If($Script:farm.offlineDatabaseSupportEnabled -eq "1")
	{
		$Script:farmofflineDatabaseSupportEnabled = $True
	}
	Else
	{
		$Script:farmofflineDatabaseSupportEnabled = $False
	}

	If($MSWord -or $PDF)
	{
		WriteWordLine 2 0 "Options"
		[System.Collections.Hashtable[]] $ScriptInformation = @()
		$ScriptInformation += @{ Data = "Auto-add"; Value = ""; }
		$ScriptInformation += @{ Data = "     Enable auto-add"; Value = $Script:FarmAutoAddEnabled.ToString(); }
		If($Script:FarmAutoAddEnabled)
		{
			$ScriptInformation += @{ Data = "     Add new devices to this site"; Value = $Script:farm.DefaultSiteName; }
		}
		$ScriptInformation += @{ Data = "Auditing"; Value = ""; }
		$ScriptInformation += @{ Data = "     Enable auditing"; Value = $Script:farmauditingEnabled.ToString(); }
		$ScriptInformation += @{ Data = "Offline database support"; Value = ""; }
		$ScriptInformation += @{ Data = "     Enable offline database support"; Value = $Script:farmofflineDatabaseSupportEnabled.ToString(); }
		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitContent;

		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If($Text)
	{
		Line 1 "Options"
		Line 2 "Auto-Add"
		Line 3 "Enable auto-add: " $Script:FarmAutoAddEnabled.ToString()
		If($Script:FarmAutoAddEnabled)
		{
			Line 3 "Add new devices to this site: " $Script:farm.DefaultSiteName
		}
		Line 2 "Auditing"
		Line 3 "Enable auditing: " $Script:farmauditingEnabled.ToString()
		Line 2 "Offline database support"
		Line 3 "Enable offline database support: " $Script:farmofflineDatabaseSupportEnabled.ToString()
		Line 0 ""
	}
	If($HTML)
	{
		WriteHTMLLine 2 0 "Options"
		$rowdata = @()
		$columnHeaders = @("Auto-add",($global:htmlsb),"",$htmlwhite)
		$rowdata += @(,('     Enable auto-add',($global:htmlsb),$Script:FarmAutoAddEnabled.ToString(),$htmlwhite))
		If($Script:FarmAutoAddEnabled)
		{
			$rowdata += @(,('     Add new devices to this site',($global:htmlsb),$Script:farm.DefaultSiteName,$htmlwhite))
		}
		$rowdata += @(,('Auditing',($global:htmlsb),"",$htmlwhite))
		$rowdata += @(,('     Enable auditing',($global:htmlsb),$Script:farmauditingEnabled.ToString(),$htmlwhite))
		$rowdata += @(,('Offline database support',($global:htmlsb),"",$htmlwhite))
		$rowdata += @(,('     Enable offline database support',($global:htmlsb),$Script:farmofflineDatabaseSupportEnabled.ToString(),$htmlwhite))
		FormatHTMLTable "" "auto" -rowArray $rowdata -columnArray $columnHeaders
	}

	#vDisk Version tab
	Write-Verbose "$(Get-Date -Format G): `tProcessing vDisk Version Tab"
	$xmergeMode = ""
	Switch ($Script:Farm.mergeMode)
	{
		0   	{$xmergeMode = "Production"; Break}
		1   	{$xmergeMode = "Test"; Break}
		2   	{$xmergeMode = "Maintenance"; Break}
		Default {$xmergeMode = "Default access mode could not be determined: $($Script:Farm.mergeMode)"; Break}
	}
	$xautomaticMergeEnabled = $farm.automaticMergeEnabled.ToString()

	If($MSWord -or $PDF)
	{
		WriteWordLine 2 0 "vDisk Version"
		[System.Collections.Hashtable[]] $ScriptInformation = @()
		$ScriptInformation += @{ Data = "Alert if number of versions from base image exceeds"; Value = $Script:farm.maxVersions.ToString(); }
		$ScriptInformation += @{ Data = "Merge after automated vDisk update, if over alert threshold"; Value = $xautomaticMergeEnabled; }
		$ScriptInformation += @{ Data = "Default access mode for new merge versions"; Value = $xmergeMode; }
		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitContent;

		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If($Text)
	{
		Line 1 "vDisk Version"
		Line 2 "Alert if number of versions from base image exceeds`t`t: " $Script:farm.maxVersions.ToString()
		Line 2 "Merge after automated vDisk update, if over alert threshold`t: " $xautomaticMergeEnabled
		Line 2 "Default access mode for new merge versions`t`t`t: " $xmergeMode
		Line 0 ""
	}
	If($HTML)
	{
		WriteHTMLLine 2 0 "vDisk Version"
		$rowdata = @()
		$columnHeaders = @("Alert if number of versions from base image exceeds",($global:htmlsb),$Script:farm.maxVersions.ToString(),$htmlwhite)
		$rowdata += @(,('Merge after automated vDisk update, if over alert threshold',($global:htmlsb),$xautomaticMergeEnabled,$htmlwhite))
		$rowdata += @(,('Default access mode for new merge versions',($global:htmlsb),$xmergeMode,$htmlwhite))
		FormatHTMLTable "" "auto" -rowArray $rowdata -columnArray $columnHeaders
	}

	#status tab
	Write-Verbose "$(Get-Date -Format G): `tProcessing Status Tab"
	$xadGroupsEnabled = ""
	If($Script:Farm.adGroupsEnabled)
	{
		$xadGroupsEnabled = "Active Directory groups are used for access rights"
	}
	Else
	{
		$xadGroupsEnabled = "Active Directory groups are not used for access rights"
	}
	
	If($Script:PVSFullVersion -ge "7.11")
	{
		$MultiSubnetFailover = $Script:farm.MultiSubnetFailover
	}
	Else
	{
		$MultiSubnetFailover = "Not supported on PVS $($Script:PVSFullVersion)"
	}
	
	If($MSWord -or $PDF)
	{
		WriteWordLine 2 0 "Status"
		WriteWordLine 0 0 "Current status of the farm:"
		[System.Collections.Hashtable[]] $ScriptInformation = @()
		$ScriptInformation += @{ Data = "Database server"; Value = $Script:farm.databaseServerName; }
		$ScriptInformation += @{ Data = "Database server IP"; Value = $SQLServerIPAddress; }
		$ScriptInformation += @{ Data = "Database instance"; Value = $Script:farm.databaseInstanceName; }
		$ScriptInformation += @{ Data = "Database"; Value = $Script:farm.databaseName; }
		$ScriptInformation += @{ Data = "Failover Partner Server"; Value = $Script:farm.failoverPartnerServerName; }
		$ScriptInformation += @{ Data = "Failover Partner Server IP"; Value = $FailoverSQLServerIPAddress; }
		$ScriptInformation += @{ Data = "Failover Partner Instance"; Value = $Script:farm.failoverPartnerServerName; }
		$ScriptInformation += @{ Data = "MultiSubnetFailover"; Value = $MultiSubnetFailover; }
		$ScriptInformation += @{ Data = $xadGroupsEnabled; Value = ""; }
		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitContent;

		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If($Text)
	{
		Line 1 "Status"
		Line 2 "Current status of the farm:"
		Line 3 "Database server`t`t`t: " $Script:farm.databaseServerName
		Line 3 "Database server IP`t`t: " $SQLServerIPAddress
		Line 3 "Database instance`t`t: " $Script:farm.databaseInstanceName
		Line 3 "Database`t`t`t: " $Script:farm.databaseName
		Line 3 "Failover Partner Server`t`t: " $Script:farm.failoverPartnerServerName
		Line 3 "Failover Partner Server IP`t: " $FailoverSQLServerIPAddress
		Line 3 "Failover Partner Instance`t: " $Script:farm.failoverPartnerInstanceName
		Line 3 "MultiSubnetFailover`t`t: " $MultiSubnetFailover
		Line 3 $xadGroupsEnabled
		Line 0 ""
	}
	If($HTML)
	{
		WriteHTMLLine 2 0 "Status"
		$rowdata = @()
		$columnHeaders = @("Database server",($global:htmlsb),$Script:farm.databaseServerName,$htmlwhite)
		$rowdata += @(,('Database server IP',($global:htmlsb),$SQLServerIPAddress,$htmlwhite))
		$rowdata += @(,('Database instance',($global:htmlsb),$Script:farm.databaseInstanceName,$htmlwhite))
		$rowdata += @(,('Database',($global:htmlsb),$Script:farm.databaseName,$htmlwhite))
		$rowdata += @(,('Failover Partner Server',($global:htmlsb),$Script:farm.failoverPartnerServerName,$htmlwhite))
		$rowdata += @(,('Failover Partner Server IP',($global:htmlsb),$FailoverSQLServerIPAddress,$htmlwhite))
		$rowdata += @(,('Failover Partner Instance',($global:htmlsb),$Script:farm.failoverPartnerInstanceName,$htmlwhite))
		$rowdata += @(,('MultiSubnetFailover',($global:htmlsb),$MultiSubnetFailover,$htmlwhite))
		$rowdata += @(,('',($global:htmlsb),$xadGroupsEnabled,$htmlwhite))
		
		$msg = "Current status of the farm"
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
	}

	#7.11 Problem Report tab
	If($Script:PVSFullVersion -ge "7.11")
	{
		Write-Verbose "$(Get-Date -Format G): `tProcessing Problem Report"
		
		$GetWhat = "cisdata"
		$GetParam = ""
		$ErrorTxt = "Problem Report information"
		$ProblemReports = BuildPVSObject $GetWhat $GetParam $ErrorTxt
		
		If($Null -ne $ProblemReports)
		{
			If($ProblemReports.UserName -eq "")
			{
				$CISUserName = "not configured"
			}
			Else
			{
				$CISUserName = $Results.UserName
			}

			$obj1 = [PSCustomObject] @{
				ItemText = "Problem report Citrix Username is $($CISUserName)"
			}
			$null = $Script:ItemsToReview.Add($obj1)

			If($MSWord -or $PDF)
			{
				WriteWordLine 2 0 "Problem Report"
				WriteWordLine 0 0 "Configure your My Citrix credentials in order to submit problem reports"
				[System.Collections.Hashtable[]] $ScriptInformation = @()
				$ScriptInformation += @{ Data = "My Citrix Username"; Value = $CISUserName; }
				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitContent;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
			}
			If($Text)
			{
				Line 1 "Problem Report"
				Line 2 "Configure your My Citrix credentials in order to submit problem reports"
				Line 2 "My Citrix Username: " $CISUserName
				Line 0 ""
			}
			If($HTML)
			{
				WriteHTMLLine 2 0 "Problem Report"
				WriteHTMLLine 0 0 "Configure your My Citrix credentials in order to submit problem reports"
				$rowdata = @()
				$columnHeaders = @("My Citrix Username",($global:htmlsb),$CISUserName,$htmlwhite)
				
				$msg = ""
				FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
			}
		}
	}
	
	Write-Verbose "$(Get-Date -Format G): "
}
#endregion

#region process PVS Site functions
Function DeviceStatus
{
	Param($xDevice)

	If($Null -eq $xDevice -or $xDevice.status -eq "" -or $xDevice.status -eq "0")
	{
		If($MSWord -or $PDF)
		{
			[System.Collections.Hashtable[]] $ScriptInformation = @()
			$ScriptInformation += @{ Data = "Target device"; Value = "Inactive"; }
			$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data,Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitContent;

			SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""
		}
		If($Text)
		{
			Line 3 "Target device: " "Inactive"
			Line 0 ""
		}
		If($HTML)
		{
			$rowdata = @()
			$columnHeaders = @("Target device",($global:htmlsb),"Inactive",$htmlwhite)

			$msg = ""
			FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
			WriteHTMLLine 0 0 " "
		}
	}
	Else
	{
		Switch ($xDevice.diskVersionAccess)
		{
			0 		{$xDevicediskVersionAccess = "Production"; Break}
			1 		{$xDevicediskVersionAccess = "Test"; Break}
			2 		{$xDevicediskVersionAccess = "Maintenance"; Break}
			3 		{$xDevicediskVersionAccess = "Personal vDisk"; Break}
			Default {$xDevicediskVersionAccess = "vDisk access type could not be determined: $($xDevice.diskVersionAccess)"; Break}
		}

		If($Script:PVSVersion -eq "7")
		{
			Switch ($xDevice.bdmBoot)
			{
				0 		{$xDevicebdmBoot = "PXE boot"; Break}
				1 		{$xDevicebdmBoot = "BDM disk"; Break}
				Default {$xDevicebdmBoot = "Boot mode could not be determined: $($xDevice.bdmBoot)"; Break}
			}
		}

		Switch ($xDevice.licenseType)
		{
			0 		{$xDevicelicenseType = "No License"; Break}
			1 		{$xDevicelicenseType = "Desktop License"; Break}
			2 		{$xDevicelicenseType = "Server License"; Break}
			5 		{$xDevicelicenseType = "OEM SmartClient License"; Break}
			6 		{$xDevicelicenseType = "XenApp License"; Break}
			7 		{$xDevicelicenseType = "XenDesktop License"; Break}
			Default {$xDevicelicenseType = "Device license type could not be determined: $($xDevice.licenseType)"; Break}
		}

		Switch ($xDevice.logLevel)
		{
			0   	{$xDevicelogLevel = "Off"; Break}
			1   	{$xDevicelogLevel = "Fatal"; Break}
			2   	{$xDevicelogLevel = "Error"; Break}
			3   	{$xDevicelogLevel = "Warning"; Break}
			4   	{$xDevicelogLevel = "Info"; Break}
			5   	{$xDevicelogLevel = "Debug"; Break}
			6   	{$xDevicelogLevel = "Trace"; Break}
			Default {$xDevicelogLevel = "Logging level could not be determined: $($xDevice.logLevel)"; Break}
		}

		If($MSWord -or $PDF)
		{
			[System.Collections.Hashtable[]] $ScriptInformation = @()
			$ScriptInformation += @{ Data = "Target device"; Value = "Active"; }
			$ScriptInformation += @{ Data = "IP Address"; Value = $xDevice.ip; }
			$ScriptInformation += @{ Data = "Server"; Value = "$($xDevice.serverName) `($($xDevice.serverIpConnection)`: $($xDevice.serverPortConnection)`)"; }
			$ScriptInformation += @{ Data = "Retries"; Value = $xDevice.status; }
			$ScriptInformation += @{ Data = "vDisk"; Value = $xDevice.diskLocatorName; }
			$ScriptInformation += @{ Data = "vDisk version"; Value = $xDevice.diskVersion; }
			$ScriptInformation += @{ Data = "vDisk name"; Value = $xDevice.diskFileName; }
			$ScriptInformation += @{ Data = "vDisk access"; Value = $xDevicediskVersionAccess; }
			If($Script:PVSVersion -eq "7")
			{
				$ScriptInformation += @{ Data = "Local write cache disk"; Value = "$($xDevice.localWriteCacheDiskSize)GB"; }
				$ScriptInformation += @{ Data = "Boot mode"; Value = $xDevicebdmBoot; }
			}
			$ScriptInformation += @{ Data = "License type"; Value = $xDevicelicenseType; }
			$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data,Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitContent;

			SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""
			
			WriteWordLine 4 0 "Logging"
			
			[System.Collections.Hashtable[]] $ScriptInformation = @()
			$ScriptInformation += @{ Data = "Logging level"; Value = $xDevicelogLevel; }

			$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data,Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitContent;

			SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""
		}
		If($Text)
		{
			Line 3 "Target device`t`t: " "Active"
			Line 3 "IP Address`t`t: " $xDevice.ip
			Line 3 "Server`t`t`t: " "$($xDevice.serverName) `($($xDevice.serverIpConnection)`: $($xDevice.serverPortConnection)`)"
			Line 3 "Retries`t`t`t: " $xDevice.status
			Line 3 "vDisk`t`t`t: " $xDevice.diskLocatorName
			Line 3 "vDisk version`t`t: " $xDevice.diskVersion
			Line 3 "vDisk name`t`t: " $xDevice.diskFileName
			Line 3 "vDisk access`t`t: " $xDevicediskVersionAccess
			If($Script:PVSVersion -eq "7")
			{
				Line 3 "Local write cache disk`t: $($xDevice.localWriteCacheDiskSize)GB"
				Line 3 "Boot mode`t`t: " $xDevicebdmBoot
			}
			Line 3 "License type`t`t: " $xDevicelicenseType
			
			Line 0 ""
			Line 2 "Logging"
			Line 3 "Logging level`t`t: " $xDevicelogLevel
			
			Line 0 ""
		}
		If($HTML)
		{
			$rowdata = @()
			$columnHeaders = @("Target device",($global:htmlsb),"Active",$htmlwhite)
			$rowdata += @(,('IP Address',($global:htmlsb),$xDevice.ip,$htmlwhite))
			$rowdata += @(,('Server',($global:htmlsb),"$($xDevice.serverName) `($($xDevice.serverIpConnection)`: $($xDevice.serverPortConnection)`)",$htmlwhite))
			$rowdata += @(,('Retries',($global:htmlsb),$xDevice.status,$htmlwhite))
			$rowdata += @(,('vDisk',($global:htmlsb),$xDevice.diskLocatorName,$htmlwhite))
			$rowdata += @(,('vDisk version',($global:htmlsb),$xDevice.diskVersion,$htmlwhite))
			$rowdata += @(,('vDisk name',($global:htmlsb),$xDevice.diskFileName,$htmlwhite))
			$rowdata += @(,('vDisk access',($global:htmlsb),$xDevicediskVersionAccess,$htmlwhite))
			If($Script:PVSVersion -eq "7")
			{
				$rowdata += @(,('Local write cache disk',($global:htmlsb),"$($xDevice.localWriteCacheDiskSize)GB",$htmlwhite))
				$rowdata += @(,('Boot mode',($global:htmlsb),$xDevicebdmBoot,$htmlwhite))
			}
			$rowdata += @(,("License type",($global:htmlsb),$xDevicelicenseType,$htmlwhite))

			$msg = ""
			FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
			WriteHTMLLine 0 0 " "
			
			$rowdata = @()
			$columnHeaders = @("Logging level",($global:htmlsb),$xDevicelogLevel,$htmlwhite)

			$msg = "Logging"
			FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
			WriteHTMLLine 0 0 " "
		}
	}
}

Function ProcessPVSSite
{
	#build site values
	Write-Verbose "$(Get-Date -Format G): Processing Sites"
	$GetWhat = "site"
	$GetParam = ""
	$ErrorTxt = "PVS Site information"
	$PVSSites = BuildPVSObject $GetWhat $GetParam $ErrorTxt
	
	If($Null -eq $PVSSites)
	{
		Write-Host -foregroundcolor Red -backgroundcolor Black "WARNING: $(Get-Date -Format G): No Sites Found"
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "No Sites Found"
		}
		If($Text)
		{
			Line 0 "No Sites Found"
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "No Sites Found"
		}
	}
	Else
	{
		ForEach($PVSSite in $PVSSites)
		{
			Write-Verbose "$(Get-Date -Format G): `tProcessing Site $($PVSSite.siteName)"
			If($MSWord -or $PDF)
			{
				$selection.InsertNewPage()
				WriteWordLine 1 0 "Site properties"
				WriteWordLine 2 0 "General"
				[System.Collections.Hashtable[]] $ScriptInformation = @()
				$ScriptInformation += @{ Data = "Site Name"; Value = $PVSSite.siteName; }
				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitContent;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
			}
			If($Text)
			{
				Line 0 "Site properties"
				Line 1 "General"
				Line 2 "Site Name: " $PVSSite.siteName
				Line 0 ""
			}
			If($HTML)
			{
				WriteHTMLLine 1 0 "Site properties"
				WriteHTMLLine 2 0 "General"
				$rowdata = @()
				$columnHeaders = @("Site Name",($global:htmlsb),$PVSSite.siteName,$htmlwhite)
				FormatHTMLTable "" "auto" -rowArray $rowdata -columnArray $columnHeaders
			}

			#security tab
			Write-Verbose "$(Get-Date -Format G): `t`tProcessing Security Tab"
			If($MSWord -or $PDF)
			{
				WriteWordLine 2 0 "Security"
				WriteWordLine 0 0 "Groups with Site Administrator access"
			}
			If($Text)
			{
				Line 1 "Security"
				Line 2 "Groups with Site Administrator access:"
			}
			If($HTML)
			{
				WriteHTMLLine 2 0 "Security"
				WriteHTMLLine 3 0 "Groups with Site Administrator access:"
			}

			$temp = $PVSSite.SiteName
			$GetWhat = "authgroup"
			$GetParam = "sitename = $temp"
			$ErrorTxt = "Groups with Site Administrator access"
			$authgroups = BuildPVSObject $GetWhat $GetParam $ErrorTxt
			If($? -and $Null -ne $AuthGroups)
			{
				OutputauthGroups $authGroups
			}
			ElseIf($? -and $Null -eq $AuthGroups)
			{
				$txt = "There are no Site Administrators defined"
				If($MSWord -or $PDF)
				{
					WriteWordLine 0 1 $txt
					WriteWordLIne 0 0 ""
				}
				If($Text)
				{
					Line 3 $txt
					Line 0 ""
				}
				If($HTML)
				{
					WriteHTMLLine 0 1 $txt
				}
			}
			Else
			{
				$txt = "Unable to retrieve Site Administrators"
				OutputWarning $txt
			}

			#MAK tab
			#MAK User and Password are encrypted

			#options tab
			Write-Verbose "$(Get-Date -Format G): `t`tProcessing Options Tab"
			If($FarmAutoAddEnabled)
			{
				If($PVSSite.DefaultCollectionName)
				{
					$xAutoAdd = $PVSSite.DefaultCollectionName
				}
				Else
				{
					$xAutoAdd = "No Default collection"
				}
			}
			Else
			{
				$xAutoAdd = "Not enabled at the Farm level"
			}

			If($MSWord -or $PDF)
			{
				WriteWordLine 2 0 "Options"
				WriteWordLine 0 0 "Auto-Add"
				[System.Collections.Hashtable[]] $ScriptInformation = @()
				$ScriptInformation += @{ Data = "Add new devices to this collection"; Value = $xAutoAdd; }
				$ScriptInformation += @{ Data = "Seconds between vDisk inventory scans"; Value = $PVSSite.InventoryFilePollingInterval.ToString(); }
				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitContent;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
			}
			If($Text)
			{
				Line 1 "Options"
				Line 2 "Auto-Add"
				Line 3 "Add new devices to this collection`t: " $xAutoAdd
				Line 3 "Seconds between vDisk inventory scans`t: " $PVSSite.InventoryFilePollingInterval.ToString()
				Line 0 ""
			}
			If($HTML)
			{
				WriteHTMLLine 2 0 "Options"
				WriteHTMLLine 0 0 "Auto-Add"
				$rowdata = @()
				$columnHeaders = @("Add new devices to this collection",($global:htmlsb),$xAutoAdd,$htmlwhite)
				$rowdata += @(,('Seconds between vDisk inventory scans',($global:htmlsb),$PVSSite.InventoryFilePollingInterval.ToString(),$htmlwhite))
				
				$msg = ""
				FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
			}

			#vDisk Update
			Write-Verbose "$(Get-Date -Format G): `t`tProcessing vDisk Update Tab"
			If($PVSSite.enableDiskUpdate -eq "1")
			{
				If($MSWord -or $PDF)
				{
					WriteWordLine 2 0 "vDisk Update"
					[System.Collections.Hashtable[]] $ScriptInformation = @()
					$ScriptInformation += @{ Data = "Enable automatic vDisk updates on this site"; Value = "Yes"; }
					$ScriptInformation += @{ Data = "Server to run vDisk updates for this site"; Value = $PVSSite.diskUpdateServerName; }
					$Table = AddWordTable -Hashtable $ScriptInformation `
					-Columns Data,Value `
					-List `
					-Format $wdTableGrid `
					-AutoFit $wdAutoFitContent;

					SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

					$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

					FindWordDocumentEnd
					$Table = $Null
					WriteWordLine 0 0 ""
				}
				If($Text)
				{
					Line 1 "vDisk Update"
					Line 2 "Enable automatic vDisk updates on this site`t: Yes"
					Line 2 "Server to run vDisk updates for this site`t: " $PVSSite.diskUpdateServerName
					Line 0 ""
				}
				If($HTML)
				{
					WriteHTMLLine 2 0 "vDisk Update"
					$rowdata = @()
					$columnHeaders = @("Enable automatic vDisk updates on this site",($global:htmlsb),"Yes",$htmlwhite)
					$rowdata += @(,('Server to run vDisk updates for this site',($global:htmlsb),$PVSSite.diskUpdateServerName,$htmlwhite))
					FormatHTMLTable "" "auto" -rowArray $rowdata -columnArray $columnHeaders
				}
			}
			Else
			{
				If($MSWord -or $PDF)
				{
					WriteWordLine 2 0 "vDisk Update"
					[System.Collections.Hashtable[]] $ScriptInformation = @()
					$ScriptInformation += @{ Data = "Enable automatic vDisk updates on this site"; Value = "No"; }
					$Table = AddWordTable -Hashtable $ScriptInformation `
					-Columns Data,Value `
					-List `
					-Format $wdTableGrid `
					-AutoFit $wdAutoFitContent;

					SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

					$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

					FindWordDocumentEnd
					$Table = $Null
					WriteWordLine 0 0 ""
				}
				If($Text)
				{
					Line 1 "vDisk Update"
					Line 2 "Enable automatic vDisk updates on this site: No"
					Line 0 ""
				}
				If($HTML)
				{
					WriteHTMLLine 2 0 "vDisk Update"
					$rowdata = @()
					$columnHeaders = @("Enable automatic vDisk updates on this site",($global:htmlsb),"No",$htmlwhite)
					FormatHTMLTable "" "auto" -rowArray $rowdata -columnArray $columnHeaders
				}
			}
			
			#process all servers in site
			Write-Verbose "$(Get-Date -Format G): `t`tProcessing Servers in Site $($PVSSite.siteName)"
			$temp = $PVSSite.SiteName
			$GetWhat = "server"
			$GetParam = "sitename = $temp"
			$ErrorTxt = "Servers for Site $temp"
			$servers = BuildPVSObject $GetWhat $GetParam $ErrorTxt
			
			If($Null -eq $servers)
			{
				Write-Host -foregroundcolor Red -backgroundcolor Black "WARNING: $(Get-Date -Format G): `t`tNo Servers Found in Site $($PVSSite.siteName)"

				If($MSWord -or $PDF)
				{
					WriteWordLine 0 0 "No Servers Found in Site $($PVSSite.siteName)"
				}
				If($Text)
				{
					Line 0 "No Servers Found in Site $($PVSSite.siteName)"
				}
				If($HTML)
				{
					WriteHTMLLine 0 0 "No Servers Found in Site $($PVSSite.siteName)"
				}

			}
			Else
			{
				If($MSWord -or $PDF)
				{
					$selection.InsertNewPage()
					WriteWordLine 2 0 "Servers"
				}
				If($Text)
				{
					Line 0 ""
					Line 1 "Servers"
				}
				If($HTML)
				{
					WriteHTMLLine 2 0 "Servers"
				}

				$FirstServer = $True
				ForEach($Server in $Servers)
				{
					#first make sure the SOAP service is running on the server
					If(VerifyPVSSOAPService $Server.serverName)
					{
						Write-Verbose "$(Get-Date -Format G): `t`t`tProcessing Server $($Server.serverName)"
						#general tab
						Write-Verbose "$(Get-Date -Format G): `t`t`t`tProcessing General Tab"

						If($Server.eventLoggingEnabled -eq "1")
						{
							$xeventLoggingEnabled = $True
						}
						Else
						{
							$xeventLoggingEnabled = $False
							$obj1 = [PSCustomObject] @{
								ItemText = "$($Server.serverName) event logging is not enabled"
							}
							$null = $Script:ItemsToReview.Add($obj1)
						}

						If($MSWord -or $PDF)
						{
							If($FirstServer -eq $False)
							{
								$Script:Selection.InsertNewPage()
							}
							WriteWordLine 3 0 $Server.serverName
							WriteWordLine 4 0 "Server Properties"
							WriteWordLine 0 0 "General"
							[System.Collections.Hashtable[]] $ScriptInformation = @()
							$ScriptInformation += @{ Data = "Name"; Value = $Server.serverName; }
							$ScriptInformation += @{ Data = "Log events to the server's Windows Event Log"; Value = $xeventLoggingEnabled.ToString(); }
							$Table = AddWordTable -Hashtable $ScriptInformation `
							-Columns Data,Value `
							-List `
							-Format $wdTableGrid `
							-AutoFit $wdAutoFitContent;

							SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

							$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

							FindWordDocumentEnd
							$Table = $Null
							WriteWordLine 0 0 ""
						}
						If($Text)
						{
							Line 2 "Server Properties"
							Line 3 "General"
							Line 4 "Name`t`t: " $Server.serverName
							Line 4 "Log events to the server's Windows Event Log: " $xeventLoggingEnabled.ToString()
							Line 0 ""
						}
						If($HTML)
						{
							WriteHTMLLine 3 0 $Server.serverName
							WriteHTMLLine 4 0 "Server Properties"
							$rowdata = @()
							$columnHeaders = @("Name",($global:htmlsb),$Server.serverName,$htmlwhite)
							$rowdata += @(,("Log events to the server's Windows Event Log",($global:htmlsb),$xeventLoggingEnabled.ToString(),$htmlwhite))
							
							$msg = "General"
							FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
							WriteHTMLLine 0 0 " "
						}
							
						Write-Verbose "$(Get-Date -Format G): `t`t`t`tProcessing Network Tab"
						$StreamingIPs = @($server.ip.split(","))
						
						ForEach($item in $StreamingIPs)
						{
							$obj1 = [PSCustomObject] @{
								ServerName = $Server.serverName
								IPAddress  = $item
							}
							$null = $Script:StreamingIPAddresses.Add($obj1)
						}

						#get all the IPv4 addresses regardless of server OS or PVS version
						If($Server.serverName -eq $env:computername)
						{
							$ServerIPs = @(Get-NetIPAddress -AddressFamily IPv4 | Where-Object {$_.ipaddress -ne "127.0.0.1"}).IPAddress
						}
						Else
						{
							$ServerIPs = @(Get-NetIPAddress -CimSession $Server.serverName -AddressFamily IPv4 | Where-Object {$_.ipaddress -ne "127.0.0.1"}).IPAddress
						}
							
						ForEach($ServerIP in $ServerIPs)
						{
							$obj1 = [PSCustomObject] @{
								serverName = $Server.serverName
								serverIP   = $ServerIP
							}
							$null = $Script:NICIPAddresses.Add( $obj1 )
						}

						If($MSWord -or $PDF)
						{
							WriteWordLine 0 0 "Network"
							[System.Collections.Hashtable[]] $ScriptInformation = @()
							If($Script:PVSVersion -eq "7")
							{
								$ScriptInformation += @{ Data = "Streaming IP addresses"; Value = $StreamingIPs[0]; }
							}
							Else
							{
								$ScriptInformation += @{ Data = "IP addresses"; Value = $StreamingIPs[0]; }
							}

							$cnt = -1
							ForEach($tmp in $StreamingIPs)
							{
								$cnt++
								If($cnt -gt 0)
								{
									$ScriptInformation += @{ Data = ""; Value = $tmp; }
								}
							}

							$ScriptInformation += @{ Data = "Ports"; Value = ""; }
							$ScriptInformation += @{ Data = "     First port"; Value = $Server.firstPort; }
							$ScriptInformation += @{ Data = "     Last port"; Value = $Server.lastPort; }
							If($Script:PVSVersion -eq "7")
							{
								$ScriptInformation += @{ Data = "Management IP"; Value = $Server.managementIp; }
							}
							$Table = AddWordTable -Hashtable $ScriptInformation `
							-Columns Data,Value `
							-List `
							-Format $wdTableGrid `
							-AutoFit $wdAutoFitContent;

							SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

							$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

							FindWordDocumentEnd
							$Table = $Null
							WriteWordLine 0 0 ""
						}
						If($Text)
						{
							Line 3 "Network"
							If($Script:PVSVersion -eq "7")
							{
								Line 4 "Streaming IP addresses`t: " $StreamingIPs[0]
							}
							Else
							{
								Line 4 "IP addresses`t: " $StreamingIPs[0]
							}

							$cnt = -1
							ForEach($tmp in $StreamingIPs)
							{
								$cnt++
								If($cnt -gt 0)
								{
									Line 6 "  " $tmp
								}
							}

							Line 4 "Ports"
							Line 5 "First port`t: " $Server.firstPort
							Line 5 "Last port`t: " $Server.lastPort
							If($Script:PVSVersion -eq "7")
							{
								Line 4 "Management IP`t`t: " $Server.managementIp
							}
							Line 0 ""
						}
						If($HTML)
						{
							$rowdata = @()
							If($Script:PVSVersion -eq "7")
							{
								$columnHeaders = @("Streaming IP addresses",($global:htmlsb),"$($StreamingIPs[0])",$htmlwhite)
							}
							Else
							{
								$columnHeaders = @("IP addresses",($global:htmlsb),"$($StreamingIPs[0])",$htmlwhite)
							}

							$cnt = -1
							ForEach($tmp in $StreamingIPs)
							{
								$cnt++
								If($cnt -gt 0)
								{
									$rowdata += @(,('',($global:htmlsb),$tmp,$htmlwhite))
								}
							}

							$rowdata += @(,('Ports',($global:htmlsb),"",$htmlwhite))
							$rowdata += @(,('     First port',($global:htmlsb),$Server.firstPort,$htmlwhite))
							$rowdata += @(,('     Last port',($global:htmlsb),$Server.lastPort,$htmlwhite))
							If($Script:PVSVersion -eq "7")
							{
								$rowdata += @(,('Management IP',($global:htmlsb),$Server.managementIp,$htmlwhite))
							}
							$msg = "Network"
							FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
						}

						#create array for appendix A
						
						Write-Verbose "$(Get-Date -Format G): `t`t`t`t`tGather Advanced server info"
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
						If($MSWord -or $PDF)
						{
							WriteWordLine 0 2 "Name: " $Server.serverName
							WriteWordLine 0 2 "Server was not processed because the server was offLine or the SOAP Service was not running"
							WriteWordLine 0 0 ""
						}
						If($Text)
						{
							Line 2 "Name: " $Server.serverName
							Line 2 "Server was not processed because the server was offLine or the SOAP Service was not running"
							Line 0 ""
						}
						If($HTML)
						{
							WriteHTMLLine 0 2 "Name: " $Server.serverName
							WriteHTMLLine 0 2 "Server was not processed because the server was offLine or the SOAP Service was not running"
							WriteHTMLLine 0 0 ""
						}
					}
					$FirstServer = $False
				}
			}

			#process all device collections in site
			Write-Verbose "$(Get-Date -Format G): `t`tProcessing all device collections in site"
			$Temp = $PVSSite.SiteName
			$GetWhat = "Collection"
			$GetParam = "siteName = $Temp"
			$ErrorTxt = "Device Collection information"
			$Collections = BuildPVSObject $GetWhat $GetParam $ErrorTxt

			If($Null -ne $Collections)
			{
				If($MSWord -or $PDF)
				{
					WriteWordLine 2 0 "Device Collections"
				}
				If($Text)
				{
					Line 0 "Device Collections"
				}
				If($HTML)
				{
					WriteHTMLLine 2 0 "Device Collections"
				}

				ForEach($Collection in $Collections)
				{
					Write-Verbose "$(Get-Date -Format G): `t`t`tProcessing Collection $($Collection.collectionName)"
					Write-Verbose "$(Get-Date -Format G): `t`t`t`tProcessing General Tab"
					If($MSWord -or $PDF)
					{
						WriteWordLine 3 0 "Device Collection Name: " $Collection.collectionName
					}
					If($Text)
					{
						Line 2 "Device Collection Name: " $Collection.collectionName
					}
					If($HTML)
					{
						WriteHTMLLine 3 0 "Device Collection Name: " $Collection.collectionName
					}

					Write-Verbose "$(Get-Date -Format G): `t`t`t`tProcessing Security Tab"
					$Temp = $Collection.collectionId
					$GetWhat = "authGroup"
					$GetParam = "collectionId = $Temp"
					$ErrorTxt = "Device Collection information"
					$AuthGroups = BuildPVSObject $GetWhat $GetParam $ErrorTxt

					$DeviceAdmins = $False
					If($Null -ne $AuthGroups)
					{
						If($MSWord -or $PDF)
						{
							WriteWordLine 0 0 "Groups with 'Device Administrator' access"
							[System.Collections.Hashtable[]] $AuthWordTable = @();
						}
						If($Text)
						{
							Line 3 "Groups with 'Device Administrator' access:"
						}
						If($HTML)
						{
							WriteHTMLLine 0 0 "Groups with 'Device Administrator' access"
							$rowdata = @()
						}

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
									If($AuthGroupUsage.role -eq "300" -and $AuthGroupUsage.Name -eq $Collection.collectionName)
									{
										$DeviceAdmins = $True
										If($MSword -or $PDF)
										{
											$WordTableRowHash = @{Name = $AuthGroup.authGroupName;}
											$AuthWordTable += $WordTableRowHash;
										}
										If($Text)
										{
											Line 8 "   " $AuthGroup.authGroupName
										}
										If($HTML)
										{
											$rowdata += @(,(
											$AuthGroup.authGroupName,$htmlwhite))
										}
									}
								}
							}
						}
								
						If($MSword -or $PDF)
						{
							If($AuthWordTable.Count -gt 0)
							{
								$Table = AddWordTable -Hashtable $AuthWordTable `
								-Columns Name `
								-Format $wdTableGrid `
								-AutoFit $wdAutoFitContent;

								SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

								$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

								FindWordDocumentEnd
								$Table = $Null
								WriteWordLine 0 0 ""
							}
						}
						If($Text)
						{
							Line 0 ""
						}
						If($HTML)
						{
							$columnHeaders = @(
							'Name',($global:htmlsb))
							
							$msg = ""
							FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
						}
					}
					If(!$DeviceAdmins)
					{
						If($MSWord -or $PDF)
						{
							WriteWordLine 0 0 "Groups with 'Device Administrator' access: None defined"
						}
						If($Text)
						{
							Line 3 "Groups with 'Device Administrator' access: None defined"
						}
						If($HTML)
						{
							WriteHTMLLine 0 0 "Groups with 'Device Administrator' access: None defined"
						}
					}

					$DeviceOperators = $False
					If($Null -ne $AuthGroups)
					{
						Line 3 ""
						If($MSWord -or $PDF)
						{
							WriteWordLine 0 0 "Groups with 'Device Operator' access:"
							[System.Collections.Hashtable[]] $AuthWordTable = @();
						}
						If($Text)
						{
							Line 3 "Groups with 'Device Operator' access:"
						}
						If($HTML)
						{
							WriteHTMLLine 0 0 "Groups with 'Device Operator' access:"
							$rowdata = @()
						}

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
									If($AuthGroupUsage.role -eq "400" -and $AuthGroupUsage.Name -eq $Collection.collectionName)
									{
										$DeviceOperators = $True
										If($MSword -or $PDF)
										{
											$WordTableRowHash = @{Name = $AuthGroup.authGroupName;}
											$AuthWordTable += $WordTableRowHash;
										}
										If($Text)
										{
											Line 7 "      " $AuthGroup.authGroupName
										}
										If($HTML)
										{
											$rowdata += @(,(
											$AuthGroup.authGroupName,$htmlwhite))
										}
									}
								}
							}
						}
								
						If($MSword -or $PDF)
						{
							If($AuthWordTable.Count -gt 0)
							{
								$Table = AddWordTable -Hashtable $AuthWordTable `
								-Columns Name `
								-Format $wdTableGrid `
								-AutoFit $wdAutoFitContent;

								SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

								$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

								FindWordDocumentEnd
								$Table = $Null
								WriteWordLine 0 0 ""
							}
						}
						If($Text)
						{
							Line 0 ""
						}
						If($HTML)
						{
							$columnHeaders = @(
							'Name',($global:htmlsb))
							
							$msg = ""
							FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
							WriteHTMLLine 0 0 ""
						}
					}
					If(!$DeviceOperators)
					{
						If($MSWord -or $PDF)
						{
							WriteWordLine 0 0 "Groups with 'Device Operator' access: None defined"
						}
						If($Text)
						{
							Line 3 "Groups with 'Device Operator' access`t : None defined"
						}
						If($HTML)
						{
							WriteHTMLLine 0 0 "Groups with 'Device Operator' access: None defined"
							WriteHTMLLine 0 0 ""
						}
					}

					Write-Verbose "$(Get-Date -Format G): `t`t`t`tProcessing Auto-Add Tab"
					If($MSWord -or $PDF)
					{
						WriteWordLine 0 0 "Auto-Add"
					}
					If($Text)
					{
						Line 3 "Auto-Add"
					}

					If($Script:FarmAutoAddEnabled)
					{
						If($Collection.autoAddZeroFill)
						{
							$autoAddZeroFill = "Yes"
						}
						Else
						{
							$autoAddZeroFill = "No"
						}
						If([String]::IsNullOrEmpty($Collection.templateDeviceName))
						{
							$TDN = "No template device"
						}
						Else
						{
							$TDN = $Collection.templateDeviceName
						}

						If($MSWord -or $PDF)
						{
							[System.Collections.Hashtable[]] $ScriptInformation = @()
							$ScriptInformation += @{ Data = "Template target device"; Value = $TDN; }
							$ScriptInformation += @{ Data = "Device Name"; Value = ""; }
							$ScriptInformation += @{ Data = "     Prefix"; Value = $Collection.autoAddPrefix; }
							$ScriptInformation += @{ Data = "     Length"; Value = $Collection.autoAddNumberLength; }
							$ScriptInformation += @{ Data = "     Zero fill"; Value = $autoAddZeroFill; }
							$ScriptInformation += @{ Data = "     Suffix"; Value = $Collection.autoAddSuffix; }
							$ScriptInformation += @{ Data = "     Last incremental #"; Value = $Collection.lastAutoAddDeviceNumber; }

							$Table = AddWordTable -Hashtable $ScriptInformation `
							-Columns Data,Value `
							-List `
							-Format $wdTableGrid `
							-AutoFit $wdAutoFitContent;

							SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

							$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

							FindWordDocumentEnd
							$Table = $Null
							WriteWordLine 0 0 ""
						}
						If($Text)
						{
							Line 4 "Template target device: " $TDN
							Line 5 "Device Name"
							Line 6 "Prefix`t`t`t: " $Collection.autoAddPrefix
							Line 6 "Length`t`t`t: " $Collection.autoAddNumberLength
							Line 6 "Zero fill`t`t: " $autoAddZeroFill
							Line 6 "Suffix`t`t`t: " $Collection.autoAddSuffix
							Line 6 "Last incremental #`t: " $Collection.lastAutoAddDeviceNumber
							Line 0 ""
						}
						If($HTML)
						{
							$rowdata = @()
							$columnHeaders = @("Template target device",($global:htmlsb),$TDN,$htmlwhite)
							$rowdata += @(,('Device Name',($global:htmlsb),"",$htmlwhite))
							$rowdata += @(,('     Prefix',($global:htmlsb),$Collection.autoAddPrefix,$htmlwhite))
							$rowdata += @(,('     Length',($global:htmlsb),$Collection.autoAddNumberLength,$htmlwhite))
							$rowdata += @(,('     Zero fill',($global:htmlsb),$autoAddZeroFill,$htmlwhite))
							$rowdata += @(,('     Suffix',($global:htmlsb),$Collection.autoAddSuffix,$htmlwhite))
							$rowdata += @(,('     Last incremental #',($global:htmlsb),$Collection.lastAutoAddDeviceNumber,$htmlwhite))
							
							$msg = "Auto-Add"
							FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
						}
					}
					Else
					{
						If($MSWord -or $PDF)
						{
							WriteWordLine 0 0 "The auto-add feature is not enabled at the PVS Farm level"
							WriteWordLine 0 0 ""
						}
						If($Text)
						{
							Line 4 "The auto-add feature is not enabled at the PVS Farm level"
							Line 0 ""
						}
						If($HTML)
						{
							WriteHTMLLine 0 0 "The auto-add feature is not enabled at the PVS Farm level"
						}
					}
					#for each collection process the first device
					Write-Verbose "$(Get-Date -Format G): `t`t`t`tProcessing the first device in each collection"
					$Temp = $Collection.collectionId
					$GetWhat = "deviceInfo"
					$GetParam = "collectionId = $Temp"
					$ErrorTxt = "Device Info information"
					$Devices = BuildPVSObject $GetWhat $GetParam $ErrorTxt
					
					If($Null -ne $Devices)
					{
						Line 0 ""
						$Device = $Devices[0]
						Write-Verbose "$(Get-Date -Format G): `t`t`t`t`tProcessing Device $($Device.deviceName)"

						If($Device.type -eq 3)
						{
							$txt = "Device with Personal vDisk Properties"
						}
						Else
						{
							$txt = "Target Device Properties"
						}

						If($MSWord -or $PDF)
						{
							WriteWordLine 2 0 $txt
						}
						If($Text)
						{
							Line 0 $txt
						}
						If($HTML)
						{
							WriteHTMLLine 2 0 $txt
						}

						Write-Verbose "$(Get-Date -Format G): `t`t`t`t`t`tProcessing General Tab"
						If($Device.type -ne "3")
						{
							Switch ($Device.type)
							{
								0 		{$DeviceType = "Production"; Break }
								1 		{$DeviceType = "Test"; Break }
								2 		{$DeviceType = "Maintenance"; Break }
								3 		{$DeviceType = "Personal vDisk"; Break }
								Default {$DeviceType = "Device type could not be determined: $($Device.type)"; Break }
							}
							Switch ($Device.bootFrom)
							{
								1 		{$DeviceBootFrom = "vDisk"; Break }
								2 		{$DeviceBootFrom = "Hard Disk"; Break }
								3 		{$DeviceBootFrom = "Floppy Disk"; Break }
								Default {$DeviceBootFrom = "Boot from could not be determined: $($Device.bootFrom)"; Break }
							}
							If($Device.enabled -eq "1")
							{
								$DeviceEnabled = "Unchecked"
							}
							Else
							{
								$DeviceEnabled = "Checked"
							}
						}
						If($Device.localDiskEnabled -eq "1")
						{
							$DevicelocalDiskEnabled = "Yes"
						}
						Else
						{
							$DevicelocalDiskEnabled = "No"
						}

						If($MSWord -or $PDF)
						{
							WriteWordLine 3 0 "General"
							[System.Collections.Hashtable[]] $ScriptInformation = @()
							$ScriptInformation += @{ Data = "Name"; Value = $Device.deviceName; }

							If($Device.type -ne "3")
							{
								$ScriptInformation += @{ Data = "Type"; Value = $DeviceType; }
								$ScriptInformation += @{ Data = "Boot from"; Value = $DeviceBootFrom; }
							}

							$ScriptInformation += @{ Data = "MAC"; Value = $Device.deviceMac; }
							$ScriptInformation += @{ Data = "Port"; Value = $Device.port; }
							
							If($Device.type -ne "3")
							{
								$ScriptInformation += @{ Data = "Class"; Value = $Device.className; }
								$ScriptInformation += @{ Data = "Disable this device"; Value = $DeviceEnabled; }
							}
							Else
							{
								$ScriptInformation += @{ Data = "vDisk"; Value = $Device.diskLocatorName; }
								$ScriptInformation += @{ Data = "Personal vDisk Drive"; Value = $Device.pvdDriveLetter; }
							}

							If($Script:Version -ge "7.12" -and $Device.XsPvsProxyUuid -ne "00000000-0000-0000-0000-000000000000")
							{
								$ScriptInformation += @{ Data = "Configured for XenServer vDisk caching"; Value = " "; }
							}
							
							$Table = AddWordTable -Hashtable $ScriptInformation `
							-Columns Data,Value `
							-List `
							-Format $wdTableGrid `
							-AutoFit $wdAutoFitContent;

							SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

							$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

							FindWordDocumentEnd
							$Table = $Null
							WriteWordLine 0 0 ""
						}
						If($Text)
						{
							Line 1 "General"
							Line 2 "Name`t`t`t: " $Device.deviceName

							If($Device.type -ne 3)
							{
								Line 2 "Type`t`t`t: " $DeviceType
								Line 2 "Boot from`t`t: " $DeviceBootFrom
							}

							Line 2 "MAC`t`t`t: " $Device.deviceMac
							Line 2 "Port`t`t`t: " $Device.port
							
							If($Device.type -ne 3)
							{
								Line 2 "Class`t`t`t: " $Device.className
								Line 2 "Disable this device`t: " $DeviceEnabled
							}
							Else
							{
								Line 2 "vDisk`t`t`t: " $Device.diskLocatorName
								Line 2 "Personal vDisk Drive`t: " $Device.pvdDriveLetter
							}

							If($Script:Version -ge "7.12" -and $Device.XsPvsProxyUuid -ne "00000000-0000-0000-0000-000000000000")
							{
								Line 2 "Configured for XenServer vDisk caching"
							}
							Line 0 ""
						}
						If($HTML)
						{
							$rowdata = @()
							$columnHeaders = @("Name",($global:htmlsb),$Device.deviceName,$htmlwhite)

							If($Device.type -ne "3")
							{
								$rowdata += @(,('Type',($global:htmlsb),$DeviceType,$htmlwhite))
								$rowdata += @(,('Boot from',($global:htmlsb),$DeviceBootFrom,$htmlwhite))
							}

							$rowdata += @(,('MAC',($global:htmlsb),$Device.deviceMac,$htmlwhite))
							$rowdata += @(,('Port',($global:htmlsb),$Device.port,$htmlwhite))
							
							If($Device.type -ne "3")
							{
								$rowdata += @(,('Class',($global:htmlsb),$Device.className,$htmlwhite))
								$rowdata += @(,('Disable this device',($global:htmlsb),$DeviceEnabled,$htmlwhite))
							}
							Else
							{
								$rowdata += @(,('vDisk',($global:htmlsb),$Device.diskLocatorName,$htmlwhite))
								$rowdata += @(,('Personal vDisk Drive',($global:htmlsb),$Device.pvdDriveLetter,$htmlwhite))
							}

							If($Script:Version -ge "7.12" -and $Device.XsPvsProxyUuid -ne "00000000-0000-0000-0000-000000000000")
							{
								$rowdata += @(,('Configured for XenServer vDisk caching',($global:htmlsb),"",$htmlwhite))
							}
						
							$msg = "General"
							FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
							WriteHTMLLine 0 0 " "
						}

						Write-Verbose "$(Get-Date -Format G): `t`t`t`t`t`tProcessing vDisks Tab"
						If($MSWord -or $PDF)
						{
							WriteWordLine 3 0 "vDisks"
						}
						If($Text)
						{
							Line 1 "vDisks"
						}

						#process all vdisks for this device
						$Temp = $Device.deviceName
						$GetWhat = "DiskInfo"
						$GetParam = "deviceName = $Temp"
						$ErrorTxt = "Device vDisk information"
						$vDisks = BuildPVSObject $GetWhat $GetParam $ErrorTxt
						If($Null -ne $vDisks)
						{
							$vDiskArray = @()
							ForEach($vDisk in $vDisks)
							{
								$vDiskArray += "$($vDisk.storeName)`\$($vDisk.diskLocatorName)"
							}
							
							If($MSWord -or $PDF)
							{
								[System.Collections.Hashtable[]] $ScriptInformation = @()
								$ScriptInformation += @{ Data = "Name"; Value = $vDiskarray[0]; }
								$cnt = -1
								ForEach($tmp in $vDiskArray)
								{
									$cnt++
									If($cnt -gt 0)
									{
										$ScriptInformation += @{ Data = ""; Value = $tmp; }
									}
								}

								$Table = AddWordTable -Hashtable $ScriptInformation `
								-Columns Data,Value `
								-List `
								-Format $wdTableGrid `
								-AutoFit $wdAutoFitContent;

								SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

								$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

								FindWordDocumentEnd
								$Table = $Null
								WriteWordLine 0 0 ""
							}
							If($Text)
							{
								Line 2 "Name: " $vDiskArray[0]
								$cnt = -1
								ForEach($tmp in $vDiskArray)
								{
									$cnt++
									If($cnt -gt 0)
									{
										Line 5 "  " $tmp
									}
								}
							}
							If($HTML)
							{
								$rowdata = @()
								$columnHeaders = @("Name",($global:htmlsb),$vDiskArray[0],$htmlwhite)
								$cnt = -1
								ForEach($tmp in $vDiskArray)
								{
									$cnt++
									If($cnt -gt 0)
									{
										$rowdata += @(,('',($global:htmlsb),$tmp,$htmlwhite))
									}
								}
						
								$msg = "vDisks"
								FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
								WriteHTMLLine 0 0 " "
							}

							If($MSWord -or $PDF)
							{
								WriteWordLine 0 0 "List local hard drive in boot menu`t: " $DevicelocalDiskEnabled
							}
							If($Text)
							{
								Line 3 "List local hard drive in boot menu: " $DevicelocalDiskEnabled
							}
							If($HTML)
							{
								WriteHTMLLine 0 0 "List local hard drive in boot menu`t: " $DevicelocalDiskEnabled
							}
							
							DeviceStatus $Device
						}
					}
					Else
					{
						If($MSWord -or $PDF)
						{
							WriteWordLine 0 0 "No Target Devices found. Device Collection is empty."
							WriteWordLine 0 0 ""
						}
						If($Text)
						{
							Line 3 "No Target Devices found. Device Collection is empty."
							Line 0 ""
						}
						If($HTML)
						{
							WriteHTMLLine 0 0 "No Target Devices found. Device Collection is empty."
							WriteHTMLLine 0 0 ""
						}
						$obj1 = [PSCustomObject] @{
							CollectionName = $Collection.collectionName
						}
						$null = $Script:EmptyDeviceCollections.Add($obj1)
					}
				}
			}
		}
	}
}

Function VerifyPVSSOAPService
{
	Param([string]$PVSServer='')
	
	Write-Verbose "$(Get-Date -Format G): `t`t`tVerifying server $($PVSServer) is online"
	If(Test-Connection -ComputerName $PVSServer -quiet -EA 0)
	{

		Write-Verbose "$(Get-Date -Format G): `t`t`tVerifying PVS SOAP Service is running on server $($PVSServer)"
		$soapserver = $Null

		$soapserver = Get-Service -ComputerName $PVSServer -EA 0 | Where-Object {$_.Name -like "soapserver"}

		If($soapserver.Status -ne "Running")
		{
			Write-Warning "The Citrix PVS Soap Server service is not Started on server $($PVSServer)"
			Write-Warning "Server $($PVSServer) cannot be processed. See message above."
			Return $False
		}
		Else
		{
			Return $True
		}
	}
	Else
	{
		Write-Warning "The server $($PVSServer) is offline or unreachable."
		Write-Warning "Server $($PVSServer) cannot be processed. See message above."
		Return $False
	}
}

#region code for hardware data
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
	# modified 29-Apr-2018 to change from Arrays to New-Object System.Collections.ArrayList
	# modified 11-Mar-2022 changed from using Get-WmiObject to Get-CimInstance

	#Get Computer info
	Write-Verbose "$(Get-Date -Format G): `t`tProcessing WMI Computer information"
	Write-Verbose "$(Get-Date -Format G): `t`t`tHardware information"
	If($MSWord -or $PDF)
	{
		WriteWordLine 3 0 "Computer Information: $($RemoteComputerName)"
		WriteWordLine 4 0 "General Computer"
	}
	If($Text)
	{
		Line 3 "Computer Information: $($RemoteComputerName)"
		Line 4 "General Computer"
	}
	If($HTML)
	{
		WriteHTMLLine 3 0 "Computer Information: $($RemoteComputerName)"
		WriteHTMLLine 4 0 "General Computer"
	}
	
	Try
	{
		If($RemoteComputerName -eq $env:computername)
		{
			$Results = Get-CimInstance -ClassName win32_computersystem -Verbose:$False
		}
		Else
		{
			$Results = Get-CimInstance -computername $RemoteComputerName -ClassName win32_computersystem -Verbose:$False
		}
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
		If($RemoteComputerName -eq $env:computername)
		{
			[string]$ComputerOS = (Get-CimInstance -ClassName Win32_OperatingSystem -EA 0 -Verbose:$False).Caption
		}
		Else
		{
			[string]$ComputerOS = (Get-CimInstance -ClassName Win32_OperatingSystem -CimSession $RemoteComputerName -EA 0 -Verbose:$False).Caption
		}

		ForEach($Item in $ComputerItems)
		{
			OutputComputerItem $Item $ComputerOS $RemoteComputerName
		}
	}
	ElseIf(!$?)
	{
		Write-Verbose "$(Get-Date -Format G): Get-CimInstance win32_computersystem failed for $($RemoteComputerName)"
		Write-Warning "Get-CimInstance win32_computersystem failed for $($RemoteComputerName)"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "Get-CimInstance win32_computersystem failed for $($RemoteComputerName)" "" $Null 0 $False $True
		}
		If($Text)
		{
			Line 5 "Get-CimInstance win32_computersystem failed for $($RemoteComputerName)"
			Line 5 ""
		}
		If($HTML)
		{
			WriteHTMLLine 0 2 "Get-CimInstance win32_computersystem failed for $($RemoteComputerName)" -option $htmlBold
		}
	}
	Else
	{
		Write-Verbose "$(Get-Date -Format G): No results Returned for Computer information"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "No results Returned for Computer information" "" $Null 0 $False $True
		}
		If($Text)
		{
			Line 5 "No results Returned for Computer information"
		}
		If($HTML)
		{
			WriteHTMLLine 0 2 "No results Returned for Computer information" -Option $htmlBold
		}
	}
	
	#Get Disk info
	Write-Verbose "$(Get-Date -Format G): `t`t`tDrive information"

	If($MSWord -or $PDF)
	{
		WriteWordLine 4 0 "Drive(s)"
	}
	If($Text)
	{
		Line 4 "Drive(s)"
	}
	If($HTML)
	{
		WriteHTMLLine 4 0 "Drive(s)"
	}

	Try
	{
		If($RemoteComputerName -eq $env:computername)
		{
			$Results = Get-CimInstance -ClassName Win32_LogicalDisk -Verbose:$False
		}
		Else
		{
			$Results = Get-CimInstance -CimSession $RemoteComputerName -ClassName Win32_LogicalDisk -Verbose:$False
		}
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
				OutputDriveItem $drive
			}
		}
	}
	ElseIf(!$?)
	{
		Write-Verbose "$(Get-Date -Format G): Get-CimInstance Win32_LogicalDisk failed for $($RemoteComputerName)"
		Write-Warning "Get-CimInstance Win32_LogicalDisk failed for $($RemoteComputerName)"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "Get-CimInstance Win32_LogicalDisk failed for $($RemoteComputerName)" "" $Null 0 $False $True
		}
		If($Text)
		{
			Line 5 "Get-CimInstance Win32_LogicalDisk failed for $($RemoteComputerName)"
		}
		If($HTML)
		{
			WriteHTMLLine 0 2 "Get-CimInstance Win32_LogicalDisk failed for $($RemoteComputerName)" -Option $htmlBold
		}
	}
	Else
	{
		Write-Verbose "$(Get-Date -Format G): No results Returned for Drive information"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "No results Returned for Drive information" "" $Null 0 $False $True
		}
		If($Text)
		{
			Line 5 "No results Returned for Drive information"
		}
		If($HTML)
		{
			WriteHTMLLine 0 2 "No results Returned for Drive information" -Option $htmlBold
		}
	}
	
	#Get CPU's and stepping
	Write-Verbose "$(Get-Date -Format G): `t`t`tProcessor information"

	If($MSWord -or $PDF)
	{
		WriteWordLine 4 0 "Processor(s)"
	}
	If($Text)
	{
		Line 4 "Processor(s)"
	}
	If($HTML)
	{
		WriteHTMLLine 4 0 "Processor(s)"
	}

	Try
	{
		If($RemoteComputerName -eq $env:computername)
		{
			$Results = Get-CimInstance -ClassName win32_Processor -Verbose:$False
		}
		Else
		{
			$Results = Get-CimInstance -computername $RemoteComputerName -ClassName win32_Processor -Verbose:$False
		}
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
			OutputProcessorItem $processor
		}
	}
	ElseIf(!$?)
	{
		Write-Verbose "$(Get-Date -Format G): Get-CimInstance win32_Processor failed for $($RemoteComputerName)"
		Write-Warning "Get-CimInstance win32_Processor failed for $($RemoteComputerName)"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "Get-CimInstance win32_Processor failed for $($RemoteComputerName)" "" $Null 0 $False $True
		}
		If($Text)
		{
			Line 5 "Get-CimInstance win32_Processor failed for $($RemoteComputerName)"
		}
		If($HTML)
		{
			WriteHTMLLine 0 2 "Get-CimInstance win32_Processor failed for $($RemoteComputerName)" -Option $htmlBold
		}
	}
	Else
	{
		Write-Verbose "$(Get-Date -Format G): No results Returned for Processor information"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "No results Returned for Processor information" "" $Null 0 $False $True
		}
		If($Text)
		{
			Line 5 "No results Returned for Processor information"
		}
		If($HTML)
		{
			WriteHTMLLine 0 2 "No results Returned for Processor information" -Option $htmlBold
		}
	}

	#Get Nics
	Write-Verbose "$(Get-Date -Format G): `t`t`tNIC information"

	If($MSWord -or $PDF)
	{
		WriteWordLine 4 0 "Network Interface(s)"
	}
	If($Text)
	{
		Line 4 "Network Interface(s)"
	}
	If($HTML)
	{
		WriteHTMLLine 4 0 "Network Interface(s)"
	}

	[bool]$GotNics = $True
	
	Try
	{
		If($RemoteComputerName -eq $env:computername)
		{
			$Results = Get-CimInstance -ClassName win32_networkadapterconfiguration -Verbose:$False
		}
		Else
		{
			$Results = Get-CimInstance -computername $RemoteComputerName -ClassName win32_networkadapterconfiguration -Verbose:$False
		}
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
					If($RemoteComputerName -eq $env:computername)
					{
						$ThisNic = Get-CimInstance -ClassName win32_networkadapter -Verbose:$False | Where-Object {$_.index -eq $nic.index}
					}
					Else
					{
						$ThisNic = Get-CimInstance -computername $RemoteComputerName -ClassName win32_networkadapter -Verbose:$False | Where-Object {$_.index -eq $nic.index}
					}
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
					Write-Warning "$(Get-Date): Error retrieving NIC information"
					Write-Verbose "$(Get-Date -Format G): Get-CimInstance win32_networkadapterconfiguration failed for $($RemoteComputerName)"
					Write-Warning "Get-CimInstance win32_networkadapterconfiguration failed for $($RemoteComputerName)"
					If($MSWORD -or $PDF)
					{
						WriteWordLine 0 2 "Error retrieving NIC information" "" $Null 0 $False $True
					}
					If($Text)
					{
						Line 5 "Error retrieving NIC information"
					}
					If($HTML)
					{
						WriteHTMLLine 0 2 "Error retrieving NIC information" -Option $htmlBold
					}
				}
				Else
				{
					Write-Verbose "$(Get-Date -Format G): No results Returned for NIC information"
					If($MSWORD -or $PDF)
					{
						WriteWordLine 0 2 "No results Returned for NIC information" "" $Null 0 $False $True
					}
					If($Text)
					{
						Line 4 "No results Returned for NIC information"
					}
					If($HTML)
					{
						WriteHTMLLine 0 2 "No results Returned for NIC information" -Option $htmlBold
					}
				}
			}
		}	
	}
	ElseIf(!$?)
	{
		Write-Warning "$(Get-Date): Error retrieving NIC configuration information"
		Write-Verbose "$(Get-Date -Format G): Get-CimInstance win32_networkadapterconfiguration failed for $($RemoteComputerName)"
		Write-Warning "Get-CimInstance win32_networkadapterconfiguration failed for $($RemoteComputerName)"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "Error retrieving NIC configuration information" "" $Null 0 $False $True
		}
		If($Text)
		{
			Line 5 "Error retrieving NIC configuration information"
		}
		If($HTML)
		{
			WriteHTMLLine 0 2 "Error retrieving NIC configuration information" -Option $htmlBold
		}
	}
	Else
	{
		Write-Verbose "$(Get-Date -Format G): No results Returned for NIC configuration information"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "No results Returned for NIC configuration information" "" $Null 0 $False $True
		}
		If($Text)
		{
			Line 5 "No results Returned for NIC configuration information"
		}
		If($HTML)
		{
			WriteHTMLLine 0 2 "No results Returned for NIC configuration information" -Option $htmlBold
		}
	}
	
	If($MSWORD -or $PDF)
	{
		WriteWordLine 0 0 ""
	}
	If($Text)
	{
		Line 0 ""
	}
	If($HTML)
	{
		WriteHTMLLine 0 0 ""
	}
}

Function OutputComputerItem
{
	Param([object]$Item, [string]$OS, [string]$RemoteComputerName)
	
	#get computer's power plan
	#https://techcommunity.microsoft.com/t5/core-infrastructure-and-security/get-the-active-power-plan-of-multiple-servers-with-powershell/ba-p/370429
	
	try 
	{

		If($RemoteComputerName -eq $env:computername)
		{
			$PowerPlan = (Get-CimInstance -ClassName Win32_PowerPlan -Namespace "root\cimv2\power" -Verbose:$False |
				Where-Object {$_.IsActive -eq $true} |
				Select-Object @{Name = "PowerPlan"; Expression = {$_.ElementName}}).PowerPlan
		}
		Else
		{
			$PowerPlan = (Get-CimInstance -CimSession $RemoteComputerName -ClassName Win32_PowerPlan -Namespace "root\cimv2\power" -Verbose:$False |
				Where-Object {$_.IsActive -eq $true} |
				Select-Object @{Name = "PowerPlan"; Expression = {$_.ElementName}}).PowerPlan
		}
	}

	catch 
	{

		$PowerPlan = $_.Exception

	}	
	
	If($MSWord -or $PDF)
	{
		$ItemInformation = New-Object System.Collections.ArrayList
		$ItemInformation.Add(@{ Data = "Manufacturer"; Value = $Item.manufacturer; }) > $Null
		$ItemInformation.Add(@{ Data = "Model"; Value = $Item.model; }) > $Null
		$ItemInformation.Add(@{ Data = "Domain"; Value = $Item.domain; }) > $Null
		$ItemInformation.Add(@{ Data = "Operating System"; Value = $OS; }) > $Null
		$ItemInformation.Add(@{ Data = "Power Plan"; Value = $PowerPlan; }) > $Null
		$ItemInformation.Add(@{ Data = "Total Ram"; Value = "$($Item.totalphysicalram) GB"; }) > $Null
		$ItemInformation.Add(@{ Data = "Physical Processors (sockets)"; Value = $Item.NumberOfProcessors; }) > $Null
		$ItemInformation.Add(@{ Data = "Logical Processors (cores w/HT)"; Value = $Item.NumberOfLogicalProcessors; }) > $Null
		$Table = AddWordTable -Hashtable $ItemInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		## Set first column format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		## IB - set column widths without recursion
		$Table.Columns.Item(1).Width = 150;
		$Table.Columns.Item(2).Width = 200;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If($Text)
	{
		Line 5 "Manufacturer`t`t`t: " $Item.manufacturer
		Line 5 "Model`t`t`t`t: " $Item.model
		Line 5 "Domain`t`t`t`t: " $Item.domain
		Line 5 "Operating System`t`t: " $OS
		Line 5 "Power Plan`t`t`t: " $PowerPlan
		Line 5 "Total Ram`t`t`t: $($Item.totalphysicalram) GB"
		Line 5 "Physical Processors (sockets)`t: " $Item.NumberOfProcessors
		Line 5 "Logical Processors (cores w/HT)`t: " $Item.NumberOfLogicalProcessors
		Line 5 ""
	}
	If($HTML)
	{
		$rowdata = @()
		$columnHeaders = @("Manufacturer",($global:htmlsb),$Item.manufacturer,$htmlwhite)
		$rowdata += @(,('Model',($global:htmlsb),$Item.model,$htmlwhite))
		$rowdata += @(,('Domain',($global:htmlsb),$Item.domain,$htmlwhite))
		$rowdata += @(,('Operating System',($global:htmlsb),$OS,$htmlwhite))
		$rowdata += @(,('Power Plan',($global:htmlsb),$PowerPlan,$htmlwhite))
		$rowdata += @(,('Total Ram',($global:htmlsb),"$($Item.totalphysicalram) GB",$htmlwhite))
		$rowdata += @(,('Physical Processors (sockets)',($global:htmlsb),$Item.NumberOfProcessors,$htmlwhite))
		$rowdata += @(,('Logical Processors (cores w/HT)',($global:htmlsb),$Item.NumberOfLogicalProcessors,$htmlwhite))

		$msg = ""
		$columnWidths = @("150px","200px")
		FormatHTMLTable $msg -rowarray $rowdata -columnArray $columnheaders -fixedWidth $columnWidths -tablewidth "350"
	}
}

Function OutputDriveItem
{
	Param([object]$Drive)
	
	$xDriveType = ""
	Switch ($drive.drivetype)
	{
		0		{$xDriveType = "Unknown"; Break}
		1		{$xDriveType = "No Root Directory"; Break}
		2		{$xDriveType = "Removable Disk"; Break}
		3		{$xDriveType = "Local Disk"; Break}
		4		{$xDriveType = "Network Drive"; Break}
		5		{$xDriveType = "Compact Disc"; Break}
		6		{$xDriveType = "RAM Disk"; Break}
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

	If($MSWORD -or $PDF)
	{
		$DriveInformation = New-Object System.Collections.ArrayList
		$DriveInformation.Add(@{ Data = "Caption"; Value = $Drive.caption; }) > $Null
		$DriveInformation.Add(@{ Data = "Size"; Value = "$($drive.drivesize) GB"; }) > $Null
		If(![String]::IsNullOrEmpty($drive.filesystem))
		{
			$DriveInformation.Add(@{ Data = "File System"; Value = $Drive.filesystem; }) > $Null
		}
		$DriveInformation.Add(@{ Data = "Free Space"; Value = "$($drive.drivefreespace) GB"; }) > $Null
		If(![String]::IsNullOrEmpty($drive.volumename))
		{
			$DriveInformation.Add(@{ Data = "Volume Name"; Value = $Drive.volumename; }) > $Null
		}
		If(![String]::IsNullOrEmpty($drive.volumedirty))
		{
			$DriveInformation.Add(@{ Data = "Volume is Dirty"; Value = $xVolumeDirty; }) > $Null
		}
		If(![String]::IsNullOrEmpty($drive.volumeserialnumber))
		{
			$DriveInformation.Add(@{ Data = "Volume Serial Number"; Value = $Drive.volumeserialnumber; }) > $Null
		}
		$DriveInformation.Add(@{ Data = "Drive Type"; Value = $xDriveType; }) > $Null
		$Table = AddWordTable -Hashtable $DriveInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		## Set first column format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells `
		-Bold `
		-BackgroundColor $wdColorGray15;

		## IB - set column widths without recursion
		$Table.Columns.Item(1).Width = 150;
		$Table.Columns.Item(2).Width = 200;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 2 ""
	}
	If($Text)
	{
		Line 5 "Caption`t`t: " $drive.caption
		Line 5 "Size`t`t: $($drive.drivesize) GB"
		If(![String]::IsNullOrEmpty($drive.filesystem))
		{
			Line 5 "File System`t: " $drive.filesystem
		}
		Line 5 "Free Space`t: $($drive.drivefreespace) GB"
		If(![String]::IsNullOrEmpty($drive.volumename))
		{
			Line 5 "Volume Name`t: " $drive.volumename
		}
		If(![String]::IsNullOrEmpty($drive.volumedirty))
		{
			Line 5 "Volume is Dirty`t: " $xVolumeDirty
		}
		If(![String]::IsNullOrEmpty($drive.volumeserialnumber))
		{
			Line 5 "Volume Serial #`t: " $drive.volumeserialnumber
		}
		Line 5 "Drive Type`t: " $xDriveType
		Line 5 ""
	}
	If($HTML)
	{
		$rowdata = @()
		$columnHeaders = @("Caption",($global:htmlsb),$Drive.caption,$htmlwhite)
		$rowdata += @(,('Size',($global:htmlsb),"$($drive.drivesize) GB",$htmlwhite))

		If(![String]::IsNullOrEmpty($drive.filesystem))
		{
			$rowdata += @(,('File System',($global:htmlsb),$Drive.filesystem,$htmlwhite))
		}
		$rowdata += @(,('Free Space',($global:htmlsb),"$($drive.drivefreespace) GB",$htmlwhite))
		If(![String]::IsNullOrEmpty($drive.volumename))
		{
			$rowdata += @(,('Volume Name',($global:htmlsb),$Drive.volumename,$htmlwhite))
		}
		If(![String]::IsNullOrEmpty($drive.volumedirty))
		{
			$rowdata += @(,('Volume is Dirty',($global:htmlsb),$xVolumeDirty,$htmlwhite))
		}
		If(![String]::IsNullOrEmpty($drive.volumeserialnumber))
		{
			$rowdata += @(,('Volume Serial Number',($global:htmlsb),$Drive.volumeserialnumber,$htmlwhite))
		}
		$rowdata += @(,('Drive Type',($global:htmlsb),$xDriveType,$htmlwhite))

		$msg = ""
		$columnWidths = @("150px","200px")
		FormatHTMLTable $msg -rowarray $rowdata -columnArray $columnheaders -fixedWidth $columnWidths -tablewidth "350"
		WriteHTMLLine 0 0 " "
	}
}

Function OutputProcessorItem
{
	Param([object]$Processor)
	
	$xAvailability = ""
	Switch ($processor.availability)
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

	If($MSWORD -or $PDF)
	{
		$ProcessorInformation = New-Object System.Collections.ArrayList
		$ProcessorInformation.Add(@{ Data = "Name"; Value = $Processor.name; }) > $Null
		$ProcessorInformation.Add(@{ Data = "Description"; Value = $Processor.description; }) > $Null
		$ProcessorInformation.Add(@{ Data = "Max Clock Speed"; Value = "$($processor.maxclockspeed) MHz"; }) > $Null
		If($processor.l2cachesize -gt 0)
		{
			$ProcessorInformation.Add(@{ Data = "L2 Cache Size"; Value = "$($processor.l2cachesize) KB"; }) > $Null
		}
		If($processor.l3cachesize -gt 0)
		{
			$ProcessorInformation.Add(@{ Data = "L3 Cache Size"; Value = "$($processor.l3cachesize) KB"; }) > $Null
		}
		If($processor.numberofcores -gt 0)
		{
			$ProcessorInformation.Add(@{ Data = "Number of Cores"; Value = $Processor.numberofcores; }) > $Null
		}
		If($processor.numberoflogicalprocessors -gt 0)
		{
			$ProcessorInformation.Add(@{ Data = "Number of Logical Processors (cores w/HT)"; Value = $Processor.numberoflogicalprocessors; }) > $Null
		}
		$ProcessorInformation.Add(@{ Data = "Availability"; Value = $xAvailability; }) > $Null
		$Table = AddWordTable -Hashtable $ProcessorInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		## Set first column format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		## IB - set column widths without recursion
		$Table.Columns.Item(1).Width = 150;
		$Table.Columns.Item(2).Width = 200;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If($Text)
	{
		Line 5 "Name`t`t`t`t: " $processor.name
		Line 5 "Description`t`t`t: " $processor.description
		Line 5 "Max Clock Speed`t`t`t: $($processor.maxclockspeed) MHz"
		If($processor.l2cachesize -gt 0)
		{
			Line 5 "L2 Cache Size`t`t`t: $($processor.l2cachesize) KB"
		}
		If($processor.l3cachesize -gt 0)
		{
			Line 5 "L3 Cache Size`t`t`t: $($processor.l3cachesize) KB"
		}
		If($processor.numberofcores -gt 0)
		{
			Line 5 "# of Cores`t`t`t: " $processor.numberofcores
		}
		If($processor.numberoflogicalprocessors -gt 0)
		{
			Line 5 "# of Logical Procs (cores w/HT)`t: " $processor.numberoflogicalprocessors
		}
		Line 5 "Availability`t`t`t: " $xAvailability
		Line 5 ""
	}
	If($HTML)
	{
		$rowdata = @()
		$columnHeaders = @("Name",($global:htmlsb),$Processor.name,$htmlwhite)
		$rowdata += @(,('Description',($global:htmlsb),$Processor.description,$htmlwhite))

		$rowdata += @(,('Max Clock Speed',($global:htmlsb),"$($processor.maxclockspeed) MHz",$htmlwhite))
		If($processor.l2cachesize -gt 0)
		{
			$rowdata += @(,('L2 Cache Size',($global:htmlsb),"$($processor.l2cachesize) KB",$htmlwhite))
		}
		If($processor.l3cachesize -gt 0)
		{
			$rowdata += @(,('L3 Cache Size',($global:htmlsb),"$($processor.l3cachesize) KB",$htmlwhite))
		}
		If($processor.numberofcores -gt 0)
		{
			$rowdata += @(,('Number of Cores',($global:htmlsb),$Processor.numberofcores,$htmlwhite))
		}
		If($processor.numberoflogicalprocessors -gt 0)
		{
			$rowdata += @(,('Number of Logical Processors (cores w/HT)',($global:htmlsb),$Processor.numberoflogicalprocessors,$htmlwhite))
		}
		$rowdata += @(,('Availability',($global:htmlsb),$xAvailability,$htmlwhite))

		$msg = ""
		$columnWidths = @("150px","200px")
		FormatHTMLTable $msg -rowarray $rowdata -columnArray $columnheaders -fixedWidth $columnWidths -tablewidth "350"
		WriteHTMLLine 0 0 " "
	}
}

Function OutputNicItem
{
	Param([object]$Nic, [object]$ThisNic, [string]$RemoteComputerName)
	
	If($RemoteComputerName -eq $env:computername)
	{
		$powerMgmt = Get-CimInstance -ClassName MSPower_DeviceEnable -Namespace "root\wmi" -Verbose:$False |
			Where-Object{$_.InstanceName -match [regex]::Escape($ThisNic.PNPDeviceID)}
	}
	Else
	{
		$powerMgmt = Get-CimInstance -CimSession $RemoteComputerName -ClassName MSPower_DeviceEnable -Namespace "root\wmi" -Verbose:$False |
			Where-Object{$_.InstanceName -match [regex]::Escape($ThisNic.PNPDeviceID)}
	}

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
		If($RemoteComputerName -eq $env:computername)
		{
			$RSSEnabled = (Get-CimInstance -ClassName MSFT_NetAdapterRssSettingData -Namespace "root\StandardCimV2" -ea 0 -Verbose:$False).Enabled
		}
		Else
		{
			$RSSEnabled = (Get-CimInstance -CimSession $RemoteComputerName -ClassName MSFT_NetAdapterRssSettingData -Namespace "root\StandardCimV2" -ea 0 -Verbose:$False).Enabled
		}

		If($RSSEnabled)
		{
			$RSSEnabled = "Enabled"
		}
		Else
		{
			$RSSEnabled = "Disabled"
		}
	}
	
	Catch
	{
		$RSSEnabled = "Unable to determine for $RemoteComputerName"
	}

	$xIPAddress = New-Object System.Collections.ArrayList
	ForEach($IPAddress in $Nic.ipaddress)
	{
		$xIPAddress.Add("$($IPAddress)") > $Null
	}

	$xIPSubnet = New-Object System.Collections.ArrayList
	ForEach($IPSubnet in $Nic.ipsubnet)
	{
		$xIPSubnet.Add("$($IPSubnet)") > $Null
	}

	If($Null -ne $nic.dnsdomainsuffixsearchorder -and $nic.dnsdomainsuffixsearchorder.length -gt 0)
	{
		$nicdnsdomainsuffixsearchorder = $nic.dnsdomainsuffixsearchorder
		$xnicdnsdomainsuffixsearchorder = New-Object System.Collections.ArrayList
		ForEach($DNSDomain in $nicdnsdomainsuffixsearchorder)
		{
			$xnicdnsdomainsuffixsearchorder.Add("$($DNSDomain)") > $Null
		}
	}
	
	If($Null -ne $nic.dnsserversearchorder -and $nic.dnsserversearchorder.length -gt 0)
	{
		$nicdnsserversearchorder = $nic.dnsserversearchorder
		$xnicdnsserversearchorder = New-Object System.Collections.ArrayList
		ForEach($DNSServer in $nicdnsserversearchorder)
		{
			$xnicdnsserversearchorder.Add("$($DNSServer)") > $Null
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

	If($MSWORD -or $PDF)
	{
		$NicInformation = New-Object System.Collections.ArrayList
		$NicInformation.Add(@{ Data = "Name"; Value = $ThisNic.Name; }) > $Null
		If($ThisNic.Name -ne $nic.description)
		{
			$NicInformation.Add(@{ Data = "Description"; Value = $Nic.description; }) > $Null
		}
		$NicInformation.Add(@{ Data = "Connection ID"; Value = $ThisNic.NetConnectionID; }) > $Null
		If(validObject $Nic Manufacturer)
		{
			$NicInformation.Add(@{ Data = "Manufacturer"; Value = $Nic.manufacturer; }) > $Null
		}
		$NicInformation.Add(@{ Data = "Availability"; Value = $xAvailability; }) > $Null
		$NicInformation.Add(@{ Data = "Allow the computer to turn off this device to save power"; Value = $PowerSaving; }) > $Null
		$NicInformation.Add(@{ Data = "Receive Side Scaling"; Value = $RSSEnabled; }) > $Null
		$NicInformation.Add(@{ Data = "Physical Address"; Value = $Nic.macaddress; }) > $Null
		If($xIPAddress.Count -gt 1)
		{
			$NicInformation.Add(@{ Data = "IP Address"; Value = $xIPAddress[0]; }) > $Null
			$NicInformation.Add(@{ Data = "Default Gateway"; Value = $Nic.Defaultipgateway; }) > $Null
			$NicInformation.Add(@{ Data = "Subnet Mask"; Value = $xIPSubnet[0]; }) > $Null
			$cnt = -1
			ForEach($tmp in $xIPAddress)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$NicInformation.Add(@{ Data = "IP Address"; Value = $tmp; }) > $Null
					$NicInformation.Add(@{ Data = "Subnet Mask"; Value = $xIPSubnet[$cnt]; }) > $Null
				}
			}
		}
		Else
		{
			$NicInformation.Add(@{ Data = "IP Address"; Value = $xIPAddress; }) > $Null
			$NicInformation.Add(@{ Data = "Default Gateway"; Value = $Nic.Defaultipgateway; }) > $Null
			$NicInformation.Add(@{ Data = "Subnet Mask"; Value = $xIPSubnet; }) > $Null
		}
		If($nic.dhcpenabled)
		{
			$DHCPLeaseObtainedDate = $nic.ConvertToDateTime($nic.dhcpleaseobtained)
			$DHCPLeaseExpiresDate = $nic.ConvertToDateTime($nic.dhcpleaseexpires)
			$NicInformation.Add(@{ Data = "DHCP Enabled"; Value = $Nic.dhcpenabled; }) > $Null
			$NicInformation.Add(@{ Data = "DHCP Lease Obtained"; Value = $dhcpleaseobtaineddate; }) > $Null
			$NicInformation.Add(@{ Data = "DHCP Lease Expires"; Value = $dhcpleaseexpiresdate; }) > $Null
			$NicInformation.Add(@{ Data = "DHCP Server"; Value = $Nic.dhcpserver; }) > $Null
		}
		If(![String]::IsNullOrEmpty($nic.dnsdomain))
		{
			$NicInformation.Add(@{ Data = "DNS Domain"; Value = $Nic.dnsdomain; }) > $Null
		}
		If($Null -ne $nic.dnsdomainsuffixsearchorder -and $nic.dnsdomainsuffixsearchorder.length -gt 0)
		{
			$NicInformation.Add(@{ Data = "DNS Search Suffixes"; Value = $xnicdnsdomainsuffixsearchorder[0]; }) > $Null
			$cnt = -1
			ForEach($tmp in $xnicdnsdomainsuffixsearchorder)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$NicInformation.Add(@{ Data = ""; Value = $tmp; }) > $Null
				}
			}
		}
		$NicInformation.Add(@{ Data = "DNS WINS Enabled"; Value = $xdnsenabledforwinsresolution; }) > $Null
		If($Null -ne $nic.dnsserversearchorder -and $nic.dnsserversearchorder.length -gt 0)
		{
			$NicInformation.Add(@{ Data = "DNS Servers"; Value = $xnicdnsserversearchorder[0]; }) > $Null
			$cnt = -1
			ForEach($tmp in $xnicdnsserversearchorder)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$NicInformation.Add(@{ Data = ""; Value = $tmp; }) > $Null
				}
			}
		}
		$NicInformation.Add(@{ Data = "NetBIOS Setting"; Value = $xTcpipNetbiosOptions; }) > $Null
		$NicInformation.Add(@{ Data = "WINS: Enabled LMHosts"; Value = $xwinsenablelmhostslookup; }) > $Null
		If(![String]::IsNullOrEmpty($nic.winshostlookupfile))
		{
			$NicInformation.Add(@{ Data = "Host Lookup File"; Value = $Nic.winshostlookupfile; }) > $Null
		}
		If(![String]::IsNullOrEmpty($nic.winsprimaryserver))
		{
			$NicInformation.Add(@{ Data = "Primary Server"; Value = $Nic.winsprimaryserver; }) > $Null
		}
		If(![String]::IsNullOrEmpty($nic.winssecondaryserver))
		{
			$NicInformation.Add(@{ Data = "Secondary Server"; Value = $Nic.winssecondaryserver; }) > $Null
		}
		If(![String]::IsNullOrEmpty($nic.winsscopeid))
		{
			$NicInformation.Add(@{ Data = "Scope ID"; Value = $Nic.winsscopeid; }) > $Null
		}
		$Table = AddWordTable -Hashtable $NicInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		## Set first column format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		## IB - set column widths without recursion
		$Table.Columns.Item(1).Width = 150;
		$Table.Columns.Item(2).Width = 200;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If($Text)
	{
		Line 5 "Name`t`t`t: " $ThisNic.Name
		If($ThisNic.Name -ne $nic.description)
		{
			Line 5 "Description`t`t: " $nic.description
		}
		Line 5 "Connection ID`t`t: " $ThisNic.NetConnectionID
		If(validObject $Nic Manufacturer)
		{
			Line 5 "Manufacturer`t`t: " $Nic.manufacturer
		}
		Line 5 "Availability`t`t: " $xAvailability
		Line 5 "Allow computer to turn "
		Line 5 "off device to save power: " $PowerSaving
		Line 5 "Physical Address`t: " $nic.macaddress
		Line 5 "Receive Side Scaling`t: " $RSSEnabled
		Line 5 "IP Address`t`t: " $xIPAddress[0]
		$cnt = -1
		ForEach($tmp in $xIPAddress)
		{
			$cnt++
			If($cnt -gt 0)
			{
				Line 8 "  " $tmp
			}
		}
		Line 5 "Default Gateway`t`t: " $Nic.Defaultipgateway
		Line 5 "Subnet Mask`t`t: " $xIPSubnet[0]
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
			Line 5 "DHCP Enabled`t`t: " $nic.dhcpenabled
			Line 5 "DHCP Lease Obtained`t: " $dhcpleaseobtaineddate
			Line 5 "DHCP Lease Expires`t: " $dhcpleaseexpiresdate
			Line 5 "DHCP Server`t`t:" $nic.dhcpserver
		}
		If(![String]::IsNullOrEmpty($nic.dnsdomain))
		{
			Line 5 "DNS Domain`t`t: " $nic.dnsdomain
		}
		If($Null -ne $nic.dnsdomainsuffixsearchorder -and $nic.dnsdomainsuffixsearchorder.length -gt 0)
		{
			[int]$x = 1
			Line 5 "DNS Search Suffixes`t: " $xnicdnsdomainsuffixsearchorder[0]
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
		Line 5 "DNS WINS Enabled`t: " $xdnsenabledforwinsresolution
		If($Null -ne $nic.dnsserversearchorder -and $nic.dnsserversearchorder.length -gt 0)
		{
			[int]$x = 1
			Line 5 "DNS Servers`t`t: " $xnicdnsserversearchorder[0]
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
		Line 5 "NetBIOS Setting`t`t: " $xTcpipNetbiosOptions
		Line 5 "WINS:"
		Line 6 "Enabled LMHosts`t: " $xwinsenablelmhostslookup
		If(![String]::IsNullOrEmpty($nic.winshostlookupfile))
		{
			Line 6 "Host Lookup File`t: " $nic.winshostlookupfile
		}
		If(![String]::IsNullOrEmpty($nic.winsprimaryserver))
		{
			Line 6 "Primary Server`t: " $nic.winsprimaryserver
		}
		If(![String]::IsNullOrEmpty($nic.winssecondaryserver))
		{
			Line 6 "Secondary Server`t: " $nic.winssecondaryserver
		}
		If(![String]::IsNullOrEmpty($nic.winsscopeid))
		{
			Line 6 "Scope ID`t`t: " $nic.winsscopeid
		}
		Line 0 ""
	}
	If($HTML)
	{
		$rowdata = @()
		$columnHeaders = @("Name",($global:htmlsb),$ThisNic.Name,$htmlwhite)
		If($ThisNic.Name -ne $nic.description)
		{
			$rowdata += @(,('Description',($global:htmlsb),$Nic.description,$htmlwhite))
		}
		$rowdata += @(,('Connection ID',($global:htmlsb),$ThisNic.NetConnectionID,$htmlwhite))
		If(validObject $Nic Manufacturer)
		{
			$rowdata += @(,('Manufacturer',($global:htmlsb),$Nic.manufacturer,$htmlwhite))
		}
		$rowdata += @(,('Availability',($global:htmlsb),$xAvailability,$htmlwhite))
		$rowdata += @(,('Allow the computer to turn off this device to save power',($global:htmlsb),$PowerSaving,$htmlwhite))
		$rowdata += @(,('Physical Address',($global:htmlsb),$Nic.macaddress,$htmlwhite))
		$rowdata += @(,('Receive Side Scaling',($global:htmlsb),$RSSEnabled,$htmlwhite))
		$rowdata += @(,('IP Address',($global:htmlsb),$xIPAddress[0],$htmlwhite))
		$cnt = -1
		ForEach($tmp in $xIPAddress)
		{
			$cnt++
			If($cnt -gt 0)
			{
				$rowdata += @(,('IP Address',($global:htmlsb),$tmp,$htmlwhite))
			}
		}
		$rowdata += @(,('Default Gateway',($global:htmlsb),$Nic.Defaultipgateway[0],$htmlwhite))
		$rowdata += @(,('Subnet Mask',($global:htmlsb),$xIPSubnet[0],$htmlwhite))
		$cnt = -1
		ForEach($tmp in $xIPSubnet)
		{
			$cnt++
			If($cnt -gt 0)
			{
				$rowdata += @(,('Subnet Mask',($global:htmlsb),$tmp,$htmlwhite))
			}
		}
		If($nic.dhcpenabled)
		{
			$DHCPLeaseObtainedDate = $nic.ConvertToDateTime($nic.dhcpleaseobtained)
			$DHCPLeaseExpiresDate = $nic.ConvertToDateTime($nic.dhcpleaseexpires)
			$rowdata += @(,('DHCP Enabled',($global:htmlsb),$Nic.dhcpenabled,$htmlwhite))
			$rowdata += @(,('DHCP Lease Obtained',($global:htmlsb),$dhcpleaseobtaineddate,$htmlwhite))
			$rowdata += @(,('DHCP Lease Expires',($global:htmlsb),$dhcpleaseexpiresdate,$htmlwhite))
			$rowdata += @(,('DHCP Server',($global:htmlsb),$Nic.dhcpserver,$htmlwhite))
		}
		If(![String]::IsNullOrEmpty($nic.dnsdomain))
		{
			$rowdata += @(,('DNS Domain',($global:htmlsb),$Nic.dnsdomain,$htmlwhite))
		}
		If($Null -ne $nic.dnsdomainsuffixsearchorder -and $nic.dnsdomainsuffixsearchorder.length -gt 0)
		{
			$rowdata += @(,('DNS Search Suffixes',($global:htmlsb),$xnicdnsdomainsuffixsearchorder[0],$htmlwhite))
			$cnt = -1
			ForEach($tmp in $xnicdnsdomainsuffixsearchorder)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$rowdata += @(,('',($global:htmlsb),$tmp,$htmlwhite))
				}
			}
		}
		$rowdata += @(,('DNS WINS Enabled',($global:htmlsb),$xdnsenabledforwinsresolution,$htmlwhite))
		If($Null -ne $nic.dnsserversearchorder -and $nic.dnsserversearchorder.length -gt 0)
		{
			$rowdata += @(,('DNS Servers',($global:htmlsb),$xnicdnsserversearchorder[0],$htmlwhite))
			$cnt = -1
			ForEach($tmp in $xnicdnsserversearchorder)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$rowdata += @(,('',($global:htmlsb),$tmp,$htmlwhite))
				}
			}
		}
		$rowdata += @(,('NetBIOS Setting',($global:htmlsb),$xTcpipNetbiosOptions,$htmlwhite))
		$rowdata += @(,('WINS: Enabled LMHosts',($global:htmlsb),$xwinsenablelmhostslookup,$htmlwhite))
		If(![String]::IsNullOrEmpty($nic.winshostlookupfile))
		{
			$rowdata += @(,('Host Lookup File',($global:htmlsb),$Nic.winshostlookupfile,$htmlwhite))
		}
		If(![String]::IsNullOrEmpty($nic.winsprimaryserver))
		{
			$rowdata += @(,('Primary Server',($global:htmlsb),$Nic.winsprimaryserver,$htmlwhite))
		}
		If(![String]::IsNullOrEmpty($nic.winssecondaryserver))
		{
			$rowdata += @(,('Secondary Server',($global:htmlsb),$Nic.winssecondaryserver,$htmlwhite))
		}
		If(![String]::IsNullOrEmpty($nic.winsscopeid))
		{
			$rowdata += @(,('Scope ID',($global:htmlsb),$Nic.winsscopeid,$htmlwhite))
		}

		$msg = ""
		$columnWidths = @("150px","200px")
		FormatHTMLTable $msg -rowarray $rowdata -columnArray $columnheaders -fixedWidth $columnWidths -tablewidth "350"
		WriteHTMLLine 0 0 " "
	}
	
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

Function GetConfigWizardInfo
{
	Param([string]$ComputerName)
	
	Write-Verbose "$(Get-Date -Format G): `t`t`t`tGather Config Wizard info"
	$DHCPServicesValue = Get-RegistryValue "HKLM:\SOFTWARE\Citrix\ProvisioningServices\Wizard" "DHCPType" $ComputerName
	$PXEServiceValue = Get-RegistryValue "HKLM:\SOFTWARE\Citrix\ProvisioningServices\Wizard" "PXEType" $ComputerName
	
	$DHCPServices = ""
	$PXEServices = ""

	Switch ($DHCPServicesValue)
	{
		1073741824	{$DHCPServices = "The service that runs on another computer"; Break}
		0			{$DHCPServices = "Microsoft DHCP"; Break}
		1			{$DHCPServices = "Provisioning Services BOOTP service"; Break}
		2			{$DHCPServices = "Other BOOTP or DHCP service"; Break}
		Default		{$DHCPServices = "Unable to determine DHCPServices: $($DHCPServicesValue)"; Break}
	}

	If($DHCPServicesValue -eq 1073741824)
	{
		Switch ($PXEServiceValue)
		{
			1073741824	{$PXEServices = "The service that runs on another computer"; Break}
			1			{$PXEServices = "Provisioning Services PXE service"; Break}	#pvs7
			0			{$PXEServices = "Provisioning Services PXE service"; Break}	#pvs6
			Default		{$PXEServices = "Unable to determine PXEServices: $($PXEServiceValue)"; Break}
		}
	}
	ElseIf($DHCPServicesValue -eq 0)
	{
		Switch ($PXEServiceValue)
		{
			1073741824	{$PXEServices = "The service that runs on another computer"; Break}
			0			{$PXEServices = "Microsoft DHCP"; Break}
			1			{$PXEServices = "Provisioning Services PXE service"; Break}
			Default		{$PXEServices = "Unable to determine PXEServices: $($PXEServiceValue)"; Break}
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
			1073741824	{$PXEServices = "The service that runs on another computer"; Break}
			0			{$PXEServices = "Provisioning Services PXE service"; Break}
			Default		{$PXEServices = "Unable to determine PXEServices: $($PXEServiceValue)"; Break}
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
	
	If($MSWord -or $PDF)
	{
		WriteWordLine 3 0 "Configuration Wizard Settings"
		[System.Collections.Hashtable[]] $ScriptInformation = @()
		$ScriptInformation += @{ Data = "DHCP Services"; Value = $DHCPServices; }
		$ScriptInformation += @{ Data = "PXE Services"; Value = $PXEServices; }
		$ScriptInformation += @{ Data = "User account"; Value = $UserAccount; }
		$ScriptInformation += @{ Data = "TFTP Option"; Value = $TFTPOption; }
		If($TFTPOptionValue -eq 1)
		{
			$ScriptInformation += @{ Data = "TFTP Bootstrap Location"; Value = $TFTPBootstrapLocation; }
		}
		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitContent;

		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If($Text)
	{
		Line 4 "Configuration Wizard Settings"
		Line 5 "DHCP Services`t`t: " $DHCPServices
		Line 5 "PXE Services`t`t: " $PXEServices
		Line 5 "User account`t`t: " $UserAccount
		Line 5 "TFTP Option`t`t: " $TFTPOption
		If($TFTPOptionValue -eq 1)
		{
			Line 5 "TFTP Bootstrap Location`t: " $TFTPBootstrapLocation
		}
		Line 0 ""
	}
	If($HTML)
	{
		WriteHTMLLine 3 0 "Configuration Wizard Settings"
		$rowdata = @()
		$columnHeaders = @("DHCP Services",($global:htmlsb),$DHCPServices,$htmlwhite)
		$rowdata += @(,('PXE Services',($global:htmlsb),$PXEServices,$htmlwhite))
		$rowdata += @(,('User account',($global:htmlsb),$UserAccount,$htmlwhite))
		$rowdata += @(,('TFTP Option',($global:htmlsb),$TFTPOption,$htmlwhite))
		If($TFTPOptionValue -eq 1)
		{
			$rowdata += @(,('TFTP Bootstrap Location',($global:htmlsb),$TFTPBootstrapLocation,$htmlwhite))
		}
		
		$msg = ""
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
	}
}

Function GetDisableTaskOffloadInfo
{
	Param([string]$ComputerName)
	
	Write-Verbose "$(Get-Date -Format G): `t`t`t`tGather TaskOffload info"
	[string]$TaskOffloadValue = Get-RegistryValue "HKLM:\SYSTEM\CurrentControlSet\Services\TCPIP\Parameters" "DisableTaskOffload" $ComputerName
	
	If($TaskOffloadValue -eq "")
	{
		$TaskOffloadValue = "Not defined"
	}
	
	$obj1 = [PSCustomObject] @{
		ServerName       = $ComputerName	
		TaskOffloadValue = $TaskOffloadValue	
	}
	$null = $Script:TaskOffloadItems.Add($obj1)
	
	If($MSWord -or $PDF)
	{
		WriteWordLine 3 0 "TaskOffload Settings"
		[System.Collections.Hashtable[]] $ScriptInformation = @()
		$ScriptInformation += @{ Data = "Value"; Value = $TaskOffloadValue; }
		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitContent;

		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If($Text)
	{
		Line 4 "TaskOffload Settings"
		Line 5 "Value: " $TaskOffloadValue
		Line 0 ""
	}
	If($HTML)
	{
		WriteHTMLLine 3 0 "TaskOffload Settings"
		$rowdata = @()
		$columnHeaders = @("Value",($global:htmlsb),$TaskOffloadValue,$htmlwhite)
		
		$msg = ""
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
	}
}

Function GetBootstrapInfo
{
	Param([object]$server)

	Write-Verbose "$(Get-Date -Format G): `t`t`t`tProcessing Bootstrap files"
	If($MSWord -or $PDF)
	{
		WriteWordLine 3 0 "Bootstrap settings"
	}
	If($Text)
	{
		Line 2 "Bootstrap settings"
	}
	If($HTML)
	{
		WriteHTMLLine 3 0 "Bootstrap settings"
	}
	Write-Verbose "$(Get-Date -Format G): `t`t`t`t`tProcessing Bootstrap files for Server $($server.servername)"
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
				If($MSWord -or $PDF)
				{
					WriteWordLine 0 2 "Server Bootstrap information could not be retrieved"
					WriteWordLine 0 2 "Error returned is " $error[0].FullyQualifiedErrorId.Split(',')[0].Trim()
					WriteWordLine 0 0 ""
				}
				If($Text)
				{
					Line 4 "Server Bootstrap information could not be retrieved"
					Line 4 "Error returned is " $error[0].FullyQualifiedErrorId.Split(',')[0].Trim()
					Line 0 ""
				}
				If($HTML)
				{
					WriteHTMLLine 0 2 "Server Bootstrap information could not be retrieved"
					WriteHTMLLine 0 2 "Error returned is " $error[0].FullyQualifiedErrorId.Split(',')[0].Trim()
					WriteHTMLLine 0 0 ""
				}
			}
		}
		If($Null -ne $ServerBootstraps)
		{
			ForEach($ServerBootstrap in $ServerBootstraps)
			{
				Write-Verbose "$(Get-Date -Format G): `t`t`t`t`tProcessing Bootstrap file $($ServerBootstrap.Bootstrapname)"
				$obj1 = [PSCustomObject] @{
					ServerName 	  = $Server.serverName				
					BootstrapName = $ServerBootstrap.Bootstrapname				
					IP1        	  = $ServerBootstrap.bootserver1_Ip				
					IP2        	  = $ServerBootstrap.bootserver2_Ip				
					IP3        	  = $ServerBootstrap.bootserver3_Ip				
					IP4        	  = $ServerBootstrap.bootserver4_Ip				
				}
				$null = $Script:BootstrapItems.Add($obj1)

				Write-Verbose "$(Get-Date -Format G): `t`t`t`t`t`tProcessing Bootstrap General Tab"

				If($MSWord -or $PDF)
				{
					WriteWordLine 4 0 "Bootstrap file: " $ServerBootstrap.Bootstrapname
					WriteWordLine 5 0 "General"
					[System.Collections.Hashtable[]] $ItemsWordTable = @();
					If( $ServerBootstrap.bootserver1_Ip -eq "0.0.0.0" -and
						$ServerBootstrap.bootserver2_Ip -eq "0.0.0.0" -and
						$ServerBootstrap.bootserver3_Ip -eq "0.0.0.0" -and
						$ServerBootstrap.bootserver4_Ip -eq "0.0.0.0")
					{
						WriteWordLine 0 0 "There are no Bootstraps defined"
					}
					Else
					{
						If($ServerBootstrap.bootserver1_Ip -ne "0.0.0.0")
						{
							If($Script:PVSVersion -eq "7")
							{
								$WordTableRowHash = @{ 
								IPAddress = $ServerBootstrap.bootserver1_Ip; 
								Port = $ServerBootstrap.bootserver1_Port; 
								SubnetMask = $ServerBootstrap.bootserver1_Netmask; 
								Gateway = $ServerBootstrap.bootserver1_Gateway;}
								$ItemsWordTable += $WordTableRowHash;
							}
							Else
							{
								$WordTableRowHash = @{ 
								IPAddress = $ServerBootstrap.bootserver1_Ip; 
								SubnetMask = $ServerBootstrap.bootserver1_Netmask; 
								Gateway = $ServerBootstrap.bootserver1_Gateway;
								Port = $ServerBootstrap.bootserver1_Port;}
								$ItemsWordTable += $WordTableRowHash;
							}
						}
						
						If($ServerBootstrap.bootserver2_Ip -ne "0.0.0.0")
						{
							If($Script:PVSVersion -eq "7")
							{
								$WordTableRowHash = @{ 
								IPAddress = $ServerBootstrap.bootserver2_Ip; 
								Port = $ServerBootstrap.bootserver2_Port; 
								SubnetMask = $ServerBootstrap.bootserver2_Netmask; 
								Gateway = $ServerBootstrap.bootserver2_Gateway;}
								$ItemsWordTable += $WordTableRowHash;
							}
							Else
							{
								$WordTableRowHash = @{ 
								IPAddress = $ServerBootstrap.bootserver2_Ip; 
								SubnetMask = $ServerBootstrap.bootserver2_Netmask; 
								Gateway = $ServerBootstrap.bootserver2_Gateway;
								Port = $ServerBootstrap.bootserver2_Port;}
								$ItemsWordTable += $WordTableRowHash;
							}
						}

						If($ServerBootstrap.bootserver3_Ip -ne "0.0.0.0")
						{
							If($Script:PVSVersion -eq "7")
							{
								$WordTableRowHash = @{ 
								IPAddress = $ServerBootstrap.bootserver3_Ip; 
								Port = $ServerBootstrap.bootserver3_Port; 
								SubnetMask = $ServerBootstrap.bootserver3_Netmask; 
								Gateway = $ServerBootstrap.bootserver3_Gateway;}
								$ItemsWordTable += $WordTableRowHash;
							}
							Else
							{
								$WordTableRowHash = @{ 
								IPAddress = $ServerBootstrap.bootserver3_Ip; 
								SubnetMask = $ServerBootstrap.bootserver3_Netmask; 
								Gateway = $ServerBootstrap.bootserver3_Gateway;
								Port = $ServerBootstrap.bootserver3_Port;}
								$ItemsWordTable += $WordTableRowHash;
							}
						}
						
						If($ServerBootstrap.bootserver4_Ip -ne "0.0.0.0")
						{
							If($Script:PVSVersion -eq "7")
							{
								$WordTableRowHash = @{ 
								IPAddress = $ServerBootstrap.bootserver4_Ip; 
								Port = $ServerBootstrap.bootserver4_Port; 
								SubnetMask = $ServerBootstrap.bootserver4_Netmask; 
								Gateway = $ServerBootstrap.bootserver4_Gateway;}
								$ItemsWordTable += $WordTableRowHash;
							}
							Else
							{
								$WordTableRowHash = @{ 
								IPAddress = $ServerBootstrap.bootserver4_Ip; 
								SubnetMask = $ServerBootstrap.bootserver4_Netmask; 
								Gateway = $ServerBootstrap.bootserver4_Gateway;
								Port = $ServerBootstrap.bootserver4_Port;}
								$ItemsWordTable += $WordTableRowHash;
							}
						}
					}

					If($ItemsWordTable.Count -gt 0)
					{
						If($Script:PVSVersion -eq "7")
						{
							$Table = AddWordTable -Hashtable $ItemsWordTable `
							-Columns IPAddress, Port, SubnetMask, Gateway `
							-Headers "IP Address", "Port", "Subnet Mask", "Gateway" `
							-AutoFit $wdAutoFitContent;
						}
						Else
						{
							$Table = AddWordTable -Hashtable $ItemsWordTable `
							-Columns IPAddress, SubnetMask, Gateway, Port `
							-Headers "IP Address", "Subnet Mask", "Gateway", "Port" `
							-AutoFit $wdAutoFitContent;
						}

						SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

						$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

						FindWordDocumentEnd
						$Table = $Null
						$ItemsWordTable = $Null
						WriteWordLine 0 0 ""
					}
				}
				If($Text)
				{
					Line 4 "Bootstrap file`t: " $ServerBootstrap.Bootstrapname
					Line 5 "General"
					
					If( $ServerBootstrap.bootserver1_Ip -eq "0.0.0.0" -and
						$ServerBootstrap.bootserver2_Ip -eq "0.0.0.0" -and
						$ServerBootstrap.bootserver3_Ip -eq "0.0.0.0" -and
						$ServerBootstrap.bootserver4_Ip -eq "0.0.0.0")
					{
						Line 6 "There are no Bootstraps defined"
					}
					Else
					{
						If($ServerBootstrap.bootserver1_Ip -ne "0.0.0.0")
						{
							If($Script:PVSVersion -eq "7")
							{
								Line 6 "IP Address`t: " $ServerBootstrap.bootserver1_Ip
								Line 6 "Port`t`t: " $ServerBootstrap.bootserver1_Port
								Line 6 "Subnet Mask`t: " $ServerBootstrap.bootserver1_Netmask
								Line 6 "Gateway`t`t: " $ServerBootstrap.bootserver1_Gateway
							}
							Else
							{
								Line 6 "IP Address`t: " $ServerBootstrap.bootserver1_Ip
								Line 6 "Subnet Mask`t: " $ServerBootstrap.bootserver1_Netmask
								Line 6 "Gateway`t`t: " $ServerBootstrap.bootserver1_Gateway
								Line 6 "Port`t`t: " $ServerBootstrap.bootserver1_Port
							}
						}
						
						If($ServerBootstrap.bootserver2_Ip -ne "0.0.0.0")
						{
							If($Script:PVSVersion -eq "7")
							{
								Line 6 "IP Address`t: " $ServerBootstrap.bootserver2_Ip
								Line 6 "Port`t`t: " $ServerBootstrap.bootserver2_Port
								Line 6 "Subnet Mask`t: " $ServerBootstrap.bootserver2_Netmask
								Line 6 "Gateway`t`t: " $ServerBootstrap.bootserver2_Gateway
							}
							Else
							{
								Line 6 "IP Address`t: " $ServerBootstrap.bootserver2_Ip
								Line 6 "Subnet Mask`t: " $ServerBootstrap.bootserver2_Netmask
								Line 6 "Gateway`t`t: " $ServerBootstrap.bootserver2_Gateway
								Line 6 "Port`t`t: " $ServerBootstrap.bootserver2_Port
							}
						}
						
						If($ServerBootstrap.bootserver3_Ip -ne "0.0.0.0")
						{
							If($Script:PVSVersion -eq "7")
							{
								Line 6 "IP Address`t: " $ServerBootstrap.bootserver3_Ip
								Line 6 "Port`t`t: " $ServerBootstrap.bootserver3_Port
								Line 6 "Subnet Mask`t: " $ServerBootstrap.bootserver3_Netmask
								Line 6 "Gateway`t`t: " $ServerBootstrap.bootserver3_Gateway
							}
							Else
							{
								Line 6 "IP Address`t: " $ServerBootstrap.bootserver3_Ip
								Line 6 "Subnet Mask`t: " $ServerBootstrap.bootserver3_Netmask
								Line 6 "Gateway`t`t: " $ServerBootstrap.bootserver3_Gateway
								Line 6 "Port`t`t: " $ServerBootstrap.bootserver3_Port
							}
						}
						
						If($ServerBootstrap.bootserver4_Ip -ne "0.0.0.0")
						{
							If($Script:PVSVersion -eq "7")
							{
								Line 6 "IP Address`t: " $ServerBootstrap.bootserver4_Ip
								Line 6 "Port`t`t: " $ServerBootstrap.bootserver4_Port
								Line 6 "Subnet Mask`t: " $ServerBootstrap.bootserver4_Netmask
								Line 6 "Gateway`t`t: " $ServerBootstrap.bootserver4_Gateway
							}
							Else
							{
								Line 6 "IP Address`t: " $ServerBootstrap.bootserver4_Ip
								Line 6 "Subnet Mask`t: " $ServerBootstrap.bootserver4_Netmask
								Line 6 "Gateway`t`t: " $ServerBootstrap.bootserver4_Gateway
								Line 6 "Port`t`t: " $ServerBootstrap.bootserver4_Port
							}
						}
					}
					Line 0 ""
				}
				If($HTML)
				{
					WriteHTMLLine 4 0 "Bootstrap file: " $ServerBootstrap.Bootstrapname
					WriteWordLine 0 0 ""
					$rowdata = @()
					If( $ServerBootstrap.bootserver1_Ip -eq "0.0.0.0" -and
						$ServerBootstrap.bootserver2_Ip -eq "0.0.0.0" -and
						$ServerBootstrap.bootserver3_Ip -eq "0.0.0.0" -and
						$ServerBootstrap.bootserver4_Ip -eq "0.0.0.0")
					{
						WriteHTMLLine 0 0 "There are no Bootstraps defined"
					}
					Else
					{
						If($ServerBootstrap.bootserver1_Ip -ne "0.0.0.0")
						{
							If($Script:PVSVersion -eq "7")
							{
								$rowdata += @(,(
								$ServerBootstrap.bootserver1_Ip,$htmlwhite,
								$ServerBootstrap.bootserver1_Port,$htmlwhite,
								$ServerBootstrap.bootserver1_Netmask,$htmlwhite,
								$ServerBootstrap.bootserver1_Gateway,$htmlwhite))
							}
							Else
							{
								$rowdata += @(,(
								$ServerBootstrap.bootserver1_Ip,$htmlwhite,
								$ServerBootstrap.bootserver1_Netmask,$htmlwhite,
								$ServerBootstrap.bootserver1_Gateway,$htmlwhite,
								$ServerBootstrap.bootserver1_Port,$htmlwhite))
							}
						}
						
						If($ServerBootstrap.bootserver2_Ip -ne "0.0.0.0")
						{
							If($Script:PVSVersion -eq "7")
							{
								$rowdata += @(,(
								$ServerBootstrap.bootserver2_Ip,$htmlwhite,
								$ServerBootstrap.bootserver2_Port,$htmlwhite,
								$ServerBootstrap.bootserver2_Netmask,$htmlwhite,
								$ServerBootstrap.bootserver2_Gateway,$htmlwhite))
							}
							Else
							{
								$rowdata += @(,(
								$ServerBootstrap.bootserver2_Ip,$htmlwhite,
								$ServerBootstrap.bootserver2_Netmask,$htmlwhite,
								$ServerBootstrap.bootserver2_Gateway,$htmlwhite,
								$ServerBootstrap.bootserver2_Port,$htmlwhite))
							}
						}

						If($ServerBootstrap.bootserver3_Ip -ne "0.0.0.0")
						{
							If($Script:PVSVersion -eq "7")
							{
								$rowdata += @(,(
								$ServerBootstrap.bootserver3_Ip,$htmlwhite,
								$ServerBootstrap.bootserver3_Port,$htmlwhite,
								$ServerBootstrap.bootserver3_Netmask,$htmlwhite,
								$ServerBootstrap.bootserver3_Gateway,$htmlwhite))
							}
							Else
							{
								$rowdata += @(,(
								$ServerBootstrap.bootserver3_Ip,$htmlwhite,
								$ServerBootstrap.bootserver3_Netmask,$htmlwhite,
								$ServerBootstrap.bootserver3_Gateway,$htmlwhite,
								$ServerBootstrap.bootserver3_Port,$htmlwhite))
							}
						}
						
						If($ServerBootstrap.bootserver4_Ip -ne "0.0.0.0")
						{
							If($Script:PVSVersion -eq "7")
							{
								$rowdata += @(,(
								$ServerBootstrap.bootserver4_Ip,$htmlwhite,
								$ServerBootstrap.bootserver4_Port,$htmlwhite,
								$ServerBootstrap.bootserver4_Netmask,$htmlwhite,
								$ServerBootstrap.bootserver4_Gateway,$htmlwhite))
							}
							Else
							{
								$rowdata += @(,(
								$ServerBootstrap.bootserver4_Ip,$htmlwhite,
								$ServerBootstrap.bootserver4_Netmask,$htmlwhite,
								$ServerBootstrap.bootserver4_Gateway,$htmlwhite,
								$ServerBootstrap.bootserver4_Port,$htmlwhite))
							}
						}
					}

					If($rowdata.Count -gt 0)
					{
						If($Script:PVSVersion -eq "7")
						{
							$columnHeaders = @(
							'IP Address',($global:htmlsb),
							'Port',($global:htmlsb),
							'Subnet Mask',($global:htmlsb),
							'Gateway',($global:htmlsb))
						}
						Else
						{
							$columnHeaders = @(
							'IP Address',($global:htmlsb),
							'Subnet Mask',($global:htmlsb),
							'Gateway',($global:htmlsb),
							'Port',($global:htmlsb))
						}
							
						$msg = "General"
						FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
						WriteHTMLLine 0 0 " "
					}
				}

				Write-Verbose "$(Get-Date -Format G): `t`t`t`t`t`tProcessing Bootstrap Options Tab"

				If($ServerBootstrap.verboseMode -eq "1")
				{
					$verboseMode = "Yes"
				}
				Else
				{
					$verboseMode = "No"
				}
				If($ServerBootstrap.interruptSafeMode -eq "1")
				{
					$interruptSafeMode = "Yes"
				}
				Else
				{
					$interruptSafeMode = "No"
				}
				If($ServerBootstrap.paeMode -eq "1")
				{
					$paeMode = "Yes"
				}
				Else
				{
					$paeMode = "No"
				}
				If($ServerBootstrap.bootFromHdOnFail -eq "1")
				{
					$bootFromHdOnFail = "Reboot to Hard Drive after $($ServerBootstrap.recoveryTime) seconds"
				}
				Else
				{
					$bootFromHdOnFail = "Restore network connection"
				}

				If($MSWord -or $PDF)
				{
					WriteWordLine 5 0 "Options"
					[System.Collections.Hashtable[]] $ScriptInformation = @()
					$ScriptInformation += @{ Data = "Verbose mode"; Value = $verboseMode; }
					$ScriptInformation += @{ Data = "Interrupt safe mode"; Value = $interruptSafeMode; }
					$ScriptInformation += @{ Data = "Advanced Memory Support"; Value = $paeMode; }
					$ScriptInformation += @{ Data = "Network recovery method"; Value = $bootFromHdOnFail; }
					$ScriptInformation += @{ Data = "Timeouts"; Value = ""; }
					$ScriptInformation += @{ Data = "     Login polling timeout"; Value = "$($ServerBootstrap.pollingTimeout) (milliseconds)"; }
					$ScriptInformation += @{ Data = "     Login general timeout"; Value = "$($ServerBootstrap.generalTimeout) (milliseconds)"; }
					$Table = AddWordTable -Hashtable $ScriptInformation `
					-Columns Data,Value `
					-List `
					-Format $wdTableGrid `
					-AutoFit $wdAutoFitContent;

					SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

					$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

					FindWordDocumentEnd
					$Table = $Null
					WriteWordLine 0 0 ""
				}
				If($Text)
				{
					Line 5 "Options"
					Line 6 "Verbose mode`t`t`t: " $verboseMode
					Line 6 "Interrupt safe mode`t`t: " $interruptSafeMode
					Line 6 "Advanced Memory Support`t`t: " $paeMode
					Line 6 "Network recovery method`t`t: " $bootFromHdOnFail
					Line 6 "Timeouts"
					Line 7 "Login polling timeout`t: " "$($ServerBootstrap.pollingTimeout) (milliseconds)"
					Line 7 "Login general timeout`t: " "$($ServerBootstrap.generalTimeout) (milliseconds)"
					Line 0 ""
				}
				If($HTML)
				{
					$rowdata = @()
					$columnHeaders = @("Verbose mode",($global:htmlsb),$verboseMode,$htmlwhite)
					$rowdata += @(,('Interrupt safe mode',($global:htmlsb),$interruptSafeMode,$htmlwhite))
					$rowdata += @(,('Advanced Memory Support',($global:htmlsb),$paeMode,$htmlwhite))
					$rowdata += @(,('Network recovery method',($global:htmlsb),$bootFromHdOnFail,$htmlwhite))
					$rowdata += @(,('Timeouts',($global:htmlsb),"",$htmlwhite))
					$rowdata += @(,('     Login polling timeout',($global:htmlsb),"$($ServerBootstrap.pollingTimeout) (milliseconds)",$htmlwhite))
					$rowdata += @(,('     Login general timeout',($global:htmlsb),"$($ServerBootstrap.generalTimeout) (milliseconds)",$htmlwhite))
					
					$msg = "Options"
					FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
					WriteHTMLLine 0 0 " "
				}
			}
		}
	}
	Else
	{
		Line 4 "No Bootstrap names available"
	}
	Line 0 ""
}

Function GetPVSServiceInfo
{
	Param([string]$ComputerName)

	Write-Verbose "$(Get-Date -Format G): `t`t`t`tProcessing PVS Services for Server $($server.servername)"
	If($RemoteComputerName -eq $env:computername)
	{
		$Services = Get-CimInstance -ClassName Win32_Service -EA 0 -Verbose:$False | `
		Where-Object {$_.DisplayName -like "Citrix PVS*"} | `
		Select-Object displayname, name, status, startmode, started, startname, state | `
		Sort-Object DisplayName
	}
	Else
	{
		$Services = Get-CimInstance -CimSession $ComputerName -ClassName Win32_Service -EA 0 -Verbose:$False | `
		Where-Object {$_.DisplayName -like "Citrix PVS*"} | `
		Select-Object displayname, name, status, startmode, started, startname, state | `
		Sort-Object DisplayName
	}
	
	If($? -and $Null -ne $Services)
	{
		ForEach($Service in $Services)
		{
			$obj1 = [PSCustomObject] @{
				ServerName     = $ComputerName
				DisplayName    = $Service.DisplayName
				ServiceName    = $Service.Name
				Status         = $Service.Status
				StartMode      = $Service.StartMode
				Started        = $Service.Started.ToString()
				StartName      = $Service.StartName
				State          = $Service.State
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
}

Function GetBadStreamingIPAddresses
{
	Param([string]$ComputerName)
	#function updated by Andrew Williamson @ Fujitsu Services to handle servers with multiple NICs
	#further optimization by Michael B. Smith
	#updated 11-Mar-2022 to handle the new array type for $Script:NICIPAddresses

	#loop through the configured streaming ip address and compare to the physical configured ip addresses
	#if a streaming ip address is not in the list of physical ip addresses, it is a bad streaming ip address
	ForEach ($Stream in ($Script:StreamingIPAddresses | Where-Object {$_.Servername -eq $ComputerName})) {
		$exists = $false
		:outerLoop ForEach ($ServerNIC in ($Script:NICIPAddresses | Where-Object {$_.Servername -eq $ComputerName})) 
		{
			ForEach ($IP in $ServerNIC.serverIP) 
			{ 
				# there could be more than one IP
				If ($Stream.IPAddress -eq $IP) 
				{
					$Exists = $true
					break :outerLoop
				}
			}
		}
		If (!$exists) 
		{
			$obj1 = [PSCustomObject] @{
				ServerName = $ComputerName			
				IPAddress  = $Stream.IPAddress			
			}
			$null = $Script:BadIPs.Add($obj1)
		}
	}
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
	    $tmp = $val.ToString()
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
	
	Write-Verbose "$(Get-Date -Format G): `t`t`t`tGather Misc Registry Key data"

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

Function GetMicrosoftHotfixes 
{
	Param([string]$ComputerName)
	
	#added V1.16 get installed Microsoft Hotfixes and Updates
	Write-Verbose "$(Get-Date -Format G): `t`t`t`tRetrieving Microsoft hotfixes and updates"
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
		#do nothing
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
		Write-Verbose "$(Get-Date -Format G): `t`t`t`tRetrieving Windows installed Roles and Features"
		$results = Get-WindowsFeature -ComputerName $ComputerName -EA 0 4> $Null
		
		If($? -and $Null -ne $results)
		{
			$WinComponents = $results | Where-Object Installed | Select-Object DisplayName,Name,FeatureType | Sort-Object DisplayName 
		
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

Function GetPVSProcessInfo
{
	Param([string]$ComputerName)
	
	#Whether or not the Inventory executable is running (Inventory.exe)
	#Whether or not the Notifier executable is running (Notifier.exe)
	#Whether or not the MgmtDaemon executable is running (MgmtDaemon.exe)
	#Whether or not the StreamProcess executable is running (StreamProcess.exe)
	
	#All four of those run within the StreamService.exe process.

	Write-Verbose "$(Get-Date -Format G): `t`t`t`tRetrieving PVS Processes for Server $($server.servername)"

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

Function GetCitrixInstalledComponents 
{
	Param([string]$ComputerName)
	
	#added V1.24 get installed Citrix components
	#code adapted from the CVAD doc script
	Write-Verbose "$(Get-Date -Format G): `t`t`t`tRetrieving Citrix installed components"
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
}

#region Process vDisks in Farm functions
Function ProcessvDisksinFarm
{
	#process all vDisks in site
	Write-Verbose "$(Get-Date -Format G): `t`tProcessing all vDisks in the Farm"
	[int]$NumberofvDisks = 0
	$GetWhat = "DiskInfo"
	$GetParam = ""
	$ErrorTxt = "Disk information"
	$Disks = BuildPVSObject $GetWhat $GetParam $ErrorTxt

	If($MSWord -or $PDF)
	{
		$selection.InsertNewPage()
		WriteWordLine 2 0 "vDisk Pool"
	}
	If($Text)
	{
		Line 1 "vDisk Pool"
	}
	If($HTML)
	{
		WriteHTMLLine 2 0 "vDisk Pool"
	}
	
	If($Null -ne $Disks)
	{
		ForEach($Disk in $Disks)
		{
			Write-Verbose "$(Get-Date -Format G): `t`t`tProcessing vDisk $($Disk.diskLocatorName)"
			Write-Verbose "$(Get-Date -Format G): `t`t`tProcessing vDisk Properties"
			Write-Verbose "$(Get-Date -Format G): `t`t`t`tProcessing General Tab"

			If($Disk.writeCacheType -eq 0)
			{
				$accessMode = "Private Image (single device, read/write access)"
			}
			Else
			{
				$accessMode = "Standard Image (multi-device, read-only access)"
				Switch ($Disk.writeCacheType)
				{
					0   	{$writeCacheType = "Private Image"; Break}
					1   	{$writeCacheType = "Cache on server"
								
								$obj1 = [PSCustomObject] @{
									StoreName = $Disk.storeName								
									SiteName  = $Disk.siteName								
									vDiskName = $Disk.diskLocatorName								
								}
								$null = $Script:CacheOnServer.Add($obj1)
								Break
							}
					3   	{$writeCacheType = "Cache in device RAM"; Break}
					4   	{$writeCacheType = "Cache on device hard disk"; Break}
					6   	{$writeCacheType = "Device RAM Disk"; Break}
					7   	{$writeCacheType = "Cache on server, persistent"; Break}
					8   	{$writeCacheType = "Cache on device hard drive persisted (NT 6.1 and later)"; Break}
					9   	{$writeCacheType = "Cache in device RAM with overflow on hard disk"; Break}
					10   	{$writeCacheType = "Private Image with Asynchronous IO"; Break} #added 1811
					11   	{$writeCacheType = "Cache on server, persistent with Asynchronous IO"; Break} #added 1811
					12   	{$writeCacheType = "Cache in device RAM with overflow on hard disk with Asynchronous IO"; Break} #added 1811
					Default {$writeCacheType = "Cache type could not be determined: $($Disk.writeCacheType)"; Break}
				}
			}
			If($Disk.adPasswordEnabled -eq "1")
			{
				$adPasswordEnabled = "Yes"
			}
			Else
			{
				$adPasswordEnabled = "No"
			}
			If($Disk.printerManagementEnabled -eq "1")
			{
				$printerManagementEnabled = "Yes"
			}
			Else
			{
				$printerManagementEnabled = "No"
			}
			If($Disk.Enabled -eq "1")
			{
				$Enabled = "Yes"
			}
			Else
			{
				$Enabled = "No"
			}
			If($Script:Version -ge "7.12")
			{
				If($Disk.ClearCacheDisabled -eq 1)
				{
					$CachedSecretsCleanup = "Yes"
				}
				Else
				{
					$CachedSecretsCleanup = "No"
				}
			}
			If($Disk.autoUpdateEnabled -eq "1")
			{
				$autoUpdateEnabled = "Yes"
			}
			Else
			{
				$autoUpdateEnabled = "No"
			}
			$DiskSize = (($Disk.diskSize/1024)/1024)

			If($MSWord -or $PDF)
			{
				WriteWordLine 3 0 $Disk.diskLocatorName
				WriteWordLine 4 0 "vDisk Properties"
				WriteWordLine 0 0 "General"
				[System.Collections.Hashtable[]] $ScriptInformation = @()
				$ScriptInformation += @{ Data = "Site"; Value = $Disk.siteName; }
				$ScriptInformation += @{ Data = "Store"; Value = $Disk.storeName; }
				$ScriptInformation += @{ Data = "Filename"; Value = $Disk.diskLocatorName; }
				$ScriptInformation += @{ Data = "Size"; Value = "$($diskSize) MB"; }
				$ScriptInformation += @{ Data = "VHD block size"; Value = "$($Disk.vhdBlockSize) KB"; }
				$ScriptInformation += @{ Data = "Access mode"; Value = $accessMode; }
				If($Disk.writeCacheType -ne 0)
				{
					$ScriptInformation += @{ Data = "Cache type"; Value = $writeCacheType; }
				}
				If($Disk.writeCacheType -ne 0 -and $Disk.writeCacheType -eq 3)
				{
					$ScriptInformation += @{ Data = "Cache size"; Value = "$($Disk.writeCacheSize) MB"; }
				}
				If($Disk.writeCacheType -ne 0 -and $Disk.writeCacheType -eq 9)
				{
					$ScriptInformation += @{ Data = "Maximum RAM size"; Value = "$($Disk.writeCacheSize) MBs"; }
				}
				If(![String]::IsNullOrEmpty($Disk.menuText))
				{
					$ScriptInformation += @{ Data = "BIOS boot menu text"; Value = $Disk.menuText; }
				}
				$ScriptInformation += @{ Data = "Enable AD machine account password management"; Value = $adPasswordEnabled; }
				$ScriptInformation += @{ Data = "Enable printer management"; Value = $printerManagementEnabled; }
				$ScriptInformation += @{ Data = "Enable streaming of this vDisk"; Value = $Enabled; }
				If($Script:PVSFullVersion -ge "7.12")
				{
					$ScriptInformation += @{ Data = "Cached secrets cleanup disabled"; Value = $CachedSecretsCleanup; }
				}
				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitContent;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
			}
			If($Text)
			{
				Line 2 $Disk.diskLocatorName
				Line 3 "vDisk Properties"
				Line 4 "General"
				Line 5 "Site`t`t: " $Disk.siteName
				Line 5 "Store`t`t: " $Disk.storeName
				Line 5 "Filename`t: " $Disk.diskLocatorName
				Line 5 "Size`t`t: " "$($diskSize) MB"
				Line 5 "VHD block size`t: " "$($Disk.vhdBlockSize) KB"
				Line 5 "Access mode`t: " $accessMode
				If($Disk.writeCacheType -ne 0)
				{
					Line 5 "Cache type`t: " $writeCacheType
				}
				If($Disk.writeCacheType -ne 0 -and $Disk.writeCacheType -eq 3)
				{
					Line 5 "Cache size`t: " "$($Disk.writeCacheSize) MB"
				}
				If($Disk.writeCacheType -ne 0 -and $Disk.writeCacheType -eq 9)
				{
					Line 5 "Maximum RAM size: " "$($Disk.writeCacheSize) MBs"
				}
				If(![String]::IsNullOrEmpty($Disk.menuText))
				{
					Line 5 "BIOS boot menu text`t`t`t: " $Disk.menuText
				}
				Line 5 "Enable AD machine acct pwd mgmt`t: " $adPasswordEnabled
				Line 5 "Enable printer management`t: " $printerManagementEnabled
				Line 5 "Enable streaming of this vDisk`t: " $Enabled
				If($Script:PVSFullVersion -ge "7.12")
				{
					Line 5 "Cached secrets cleanup disabled`t: " $CachedSecretsCleanup
				}
				Line 0 ""
			}
			If($HTML)
			{
				WriteHTMLLine 3 0 $Disk.diskLocatorName
				WriteHTMLLine 4 0 "vDisk Properties"
				$rowdata = @()
				$columnHeaders = @("Site",($global:htmlsb),$Disk.siteName,$htmlwhite)
				$rowdata += @(,('Store',($global:htmlsb),$Disk.storeName,$htmlwhite))
				$rowdata += @(,('Filename',($global:htmlsb),$Disk.diskLocatorName,$htmlwhite))
				$rowdata += @(,('Size',($global:htmlsb),"$($diskSize) MB",$htmlwhite))
				$rowdata += @(,('VHD block size',($global:htmlsb),"$($Disk.vhdBlockSize) KB",$htmlwhite))
				$rowdata += @(,('Access mode',($global:htmlsb),$accessMode,$htmlwhite))
				If($Disk.writeCacheType -ne 0)
				{
					$rowdata += @(,('Cache type',($global:htmlsb),$writeCacheType,$htmlwhite))
				}
				If($Disk.writeCacheType -ne 0 -and $Disk.writeCacheType -eq 3)
				{
					$rowdata += @(,('Cache size',($global:htmlsb),"$($Disk.writeCacheSize) MB",$htmlwhite))
				}
				If($Disk.writeCacheType -ne 0 -and $Disk.writeCacheType -eq 9)
				{
					$rowdata += @(,('Maximum RAM size',($global:htmlsb),"$($Disk.writeCacheSize) MBs",$htmlwhite))
				}
				If(![String]::IsNullOrEmpty($Disk.menuText))
				{
					$rowdata += @(,('BIOS boot menu text',($global:htmlsb),$Disk.menuText,$htmlwhite))
				}
				$rowdata += @(,('Enable AD machine account password management',($global:htmlsb),$adPasswordEnabled,$htmlwhite))
				$rowdata += @(,('Enable printer management',($global:htmlsb),$printerManagementEnabled,$htmlwhite))
				$rowdata += @(,('Enable streaming of this vDisk',($global:htmlsb),$Enabled,$htmlwhite))
				If($Script:PVSFullVersion -ge "7.12")
				{
					$rowdata += @(,('Cached secrets cleanup disabled',($global:htmlsb),$CachedSecretsCleanup,$htmlwhite))
				}
				
				$msg = "General"
				FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
				WriteHTMLLine 0 0 " "
			}

			Write-Verbose "$(Get-Date -Format G): `t`t`t`tProcessing Auto Update Tab"

			If($Disk.autoUpdateEnabled -eq "1")
			{
				$autoUpdateEnabled = "Yes"
			}
			Else
			{
				$autoUpdateEnabled = "No"
			}

			If($MSWord -or $PDF)
			{
				WriteWordLine 4 0 "Auto Update"
				[System.Collections.Hashtable[]] $ScriptInformation = @()
				$ScriptInformation += @{ Data = "Enable automatic updates for the vDisk"; Value = $autoUpdateEnabled; }
				If($Disk.autoUpdateEnabled -eq "1")
				{
					If($Disk.activationDateEnabled -eq "0")
					{
						$ScriptInformation += @{ Data = "Apply vDisk updates as soon as they are detected by the server"; Value = ""; }
					}
					Else
					{
						$ScriptInformation += @{ Data = "Schedule the next vDisk update to occur on"; Value = $Disk.activeDate; }
					}
				}
				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitContent;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
			}
			If($Text)
			{
				Line 4 "Auto Update"
				Line 5 "Enable automatic updates for the vDisk: " $autoUpdateEnabled
				If($Disk.autoUpdateEnabled -eq "1")
				{
					If($Disk.activationDateEnabled -eq "0")
					{
						Line 5 "Apply vDisk updates as soon as they are detected by the server"
					}
					Else
					{
						Line 5 "Schedule the next vDisk update to occur on`t: $($Disk.activeDate)"
					}
				}
				Line 0 ""
			}
			If($HTML)
			{
				$rowdata = @()
				$columnHeaders = @("Enable automatic updates for the vDisk",($global:htmlsb),$autoUpdateEnabled,$htmlwhite)
				If($Disk.autoUpdateEnabled -eq "1")
				{
					If($Disk.activationDateEnabled -eq "0")
					{
						$rowdata += @(,('',($global:htmlsb),"Apply vDisk updates as soon as they are detected by the server",$htmlwhite))
					}
					Else
					{
						$rowdata += @(,('',($global:htmlsb),"Schedule the next vDisk update to occur on: $($Disk.activeDate)",$htmlwhite))
					}
				}
				
				$msg = "Auto Update"
				FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
			}
			
			#process Versions menu
			#get versions info
			#thanks to the PVS Product team for their help in understanding the Versions information
			Write-Verbose "$(Get-Date -Format G): `t`t`tProcessing vDisk Versions"
			If($MSWord -or $PDF)
			{
				WriteWordLine 4 0 "vDisk Versions"
			}
			If($Text)
			{
				Line 3 "vDisk Versions"
			}
			If($HTML)
			{
				WriteHTMLLine 4 0 "vDisk Versions"
			}
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
					
					$tmp = ""
					If($BootOverride)
					{
						$tmp = $BootingVersion
					}
					Else
					{
						$tmp = "Newest released"
					}

					If($MSWord -or $PDF)
					{
						WriteWordLine 0 0 "Boot production devices from version: " $tmp
					}
					If($Text)
					{
						Line 4 "Boot production devices from version`t: " $tmp
					}
					If($HTML)
					{
						WriteHTMLLine 0 0 "Boot production devices from version: " $tmp
					}
					
					$VersionFlag = $False
					ForEach($DiskVersion in $DiskVersions)
					{
						Write-Verbose "$(Get-Date -Format G): `t`t`t`tProcessing vDisk Version $($DiskVersion.version)"

						If($DiskVersion.version -gt $Script:farm.maxVersions -and $VersionFlag -eq $False)
						{
							$VersionFlag = $True
							
							$obj1 = [PSCustomObject] @{
								vDiskName = $Disk.diskLocatorName								
							}
							$null = $Script:VersionsToMerge.Add($obj1)
						}

						If($DiskVersion.version -eq $BootingVersion)
						{
							$BootFromVersion = "$($DiskVersion.version) (Current booting version)"
						}
						Else
						{
							$BootFromVersion = $DiskVersion.version.ToString()
						}

						Switch ($DiskVersion.access)
						{
							"0" 	{$access = "Production"; Break }
							"1" 	{$access = "Maintenance"; Break }
							"2" 	{$access = "Maintenance Highest Version"; Break }
							"3" 	{$access = "Override"; Break }
							"4" 	{$access = "Merge"; Break }
							"5" 	{$access = "Merge Maintenance"; Break }
							"6" 	{$access = "Merge Test"; Break }
							"7" 	{$access = "Test"; Break }
							Default {$access = "Access could not be determined: $($DiskVersion.access)"; Break }
						}

						Switch ($DiskVersion.type)
						{
							"0" 	{$DiskVersionType = "Base"; Break }
							"1" 	{$DiskVersionType = "Manual"; Break }
							"2" 	{$DiskVersionType = "Automatic"; Break }
							"3" 	{$DiskVersionType = "Merge"; Break }
							"4" 	{$DiskVersionType = "Merge Base"; Break }
							Default {$DiskVersionType = "Type could not be determined: $($DiskVersion.type)"; Break }
						}

						Switch ($DiskVersion.canDelete)
						{
							"0"	{$canDelete = "No"; Break }
							"1"	{$canDelete = "Yes"; Break }
						}

						Switch ($DiskVersion.canMerge)
						{
							"0"	{$canMerge = "No"; Break }
							"1"	{$canMerge = "Yes"; Break }
						}

						Switch ($DiskVersion.canMergeBase)
						{
							"0"	{$canMergeBase = "No"; Break }
							"1"	{$canMergeBase = "Yes"; Break }
						}

						Switch ($DiskVersion.canPromote)
						{
							"0"	{$canPromote = "No"; Break }
							"1"	{$canPromote = "Yes"; Break }
						}

						Switch ($DiskVersion.canRevertTest)
						{
							"0"	{$canRevertTest = "No"; Break }
							"1"	{$canRevertTest = "Yes"; Break }
						}

						Switch ($DiskVersion.canRevertMaintenance)
						{
							"0"	{$canRevertMaintenance = "No"; Break }
							"1"	{$canRevertMaintenance = "Yes"; Break }
						}

						Switch ($DiskVersion.canSetScheduledDate)
						{
							"0"	{$canSetScheduledDate = "No"; Break }
							"1"	{$canSetScheduledDate = "Yes"; Break }
						}

						Switch ($DiskVersion.canOverride)
						{
							"0"	{$canOverride = "No"; Break }
							"1"	{$canOverride = "Yes"; Break }
						}

						Switch ($DiskVersion.isPending)
						{
							"0"	{$isPending = "No, version Scheduled Date has occurred"; Break }
							"1"	{$isPending = "Yes, version Scheduled Date has not occurred"; Break }
						}

						Switch ($DiskVersion.goodInventoryStatus)
						{
							"0"		{$goodInventoryStatus = "Not available on all servers"; Break }
							"1"		{$goodInventoryStatus = "Available on all servers"; Break }
							Default {$goodInventoryStatus = "Replication status could not be determined: $($DiskVersion.goodInventoryStatus)"; Break }
						}

						If($MSWord -or $PDF)
						{
							[System.Collections.Hashtable[]] $ScriptInformation = @()
							$ScriptInformation += @{ Data = "Version"; Value = $BootFromVersion; }
							If($DiskVersion.version -gt $Script:farm.maxVersions -and $VersionFlag -eq $False)
							{
								$ScriptInformation += @{ Data = "Version of vDisk is $($DiskVersion.version) which is greater than the limit of $($Script:farm.maxVersions)"; Value = "Consider merging"; }
							}
							$ScriptInformation += @{ Data = "Created"; Value = $DiskVersion.createDate; }
							If(![String]::IsNullOrEmpty($DiskVersion.scheduledDate))
							{
								$ScriptInformation += @{ Data = "Released"; Value = $DiskVersion.scheduledDate; }
							}
							$ScriptInformation += @{ Data = "Devices"; Value = $DiskVersion.deviceCount; }
							$ScriptInformation += @{ Data = "Access"; Value = $access; }
							$ScriptInformation += @{ Data = "Type"; Value = $DiskVersionType; }
							If(![String]::IsNullOrEmpty($DiskVersion.description))
							{
								$ScriptInformation += @{ Data = "Properties"; Value = $DiskVersion.description; }
							}
							$ScriptInformation += @{ Data = "Can Delete"; Value = $canDelete; }
							$ScriptInformation += @{ Data = "Can Merge"; Value = $canMerge; }
							$ScriptInformation += @{ Data = "Can Merge Base"; Value = $canMergeBase; }
							$ScriptInformation += @{ Data = "Can Promote"; Value = $canPromote; }
							$ScriptInformation += @{ Data = "Can Revert back to Test"; Value = $canRevertTest; }
							$ScriptInformation += @{ Data = "Can Revert back to Maintenance"; Value = $canRevertMaintenance; }
							$ScriptInformation += @{ Data = "Can Set Scheduled Date"; Value = $canSetScheduledDate; }
							$ScriptInformation += @{ Data = "Can Override"; Value = $canOverride; }
							$ScriptInformation += @{ Data = "Is Pending"; Value = $isPending; }
							$ScriptInformation += @{ Data = "Replication Status"; Value = $goodInventoryStatus; }
							$ScriptInformation += @{ Data = "Disk Filename"; Value = $DiskVersion.diskFileName; }
							$Table = AddWordTable -Hashtable $ScriptInformation `
							-Columns Data,Value `
							-List `
							-Format $wdTableGrid `
							-AutoFit $wdAutoFitContent;

							SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

							$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

							FindWordDocumentEnd
							$Table = $Null
							WriteWordLine 0 0 ""
						}
						If($Text)
						{
							Line 4 "Version`t`t`t`t`t: " $BootFromVersion
							If($DiskVersion.version -gt $Script:farm.maxVersions -and $VersionFlag -eq $False)
							{
								Line 4 "Version of vDisk is $($DiskVersion.version) which is greater than the limit of $($Script:farm.maxVersions). Consider merging."
							}
							Line 4 "Created`t`t`t`t`t: " $DiskVersion.createDate
							If(![String]::IsNullOrEmpty($DiskVersion.scheduledDate))
							{
								Line 4 "Released`t`t`t`t: " $DiskVersion.scheduledDate
							}
							Line 4 "Devices`t`t`t`t`t: " $DiskVersion.deviceCount
							Line 4 "Access`t`t`t`t`t: " $access
							Line 4 "Type`t`t`t`t`t: " $DiskVersionType
							If(![String]::IsNullOrEmpty($DiskVersion.description))
							{
								Line 4 "Properties`t`t`t`t: " $DiskVersion.description
							}
							Line 4 "Can Delete`t`t`t`t: " $canDelete
							Line 4 "Can Merge`t`t`t`t: " $canMerge
							Line 4 "Can Merge Base`t`t`t`t: " $canMergeBase
							Line 4 "Can Promote`t`t`t`t: " $canPromote
							Line 4 "Can Revert back to Test`t`t`t: " $canRevertTest
							Line 4 "Can Revert back to Maintenance`t`t: " $canRevertMaintenance
							Line 4 "Can Set Scheduled Date`t`t`t: " $canSetScheduledDate
							Line 4 "Can Override`t`t`t`t: " $canOverride
							Line 4 "Is Pending`t`t`t`t: " $isPending
							Line 4 "Replication Status`t`t`t: " $goodInventoryStatus
							Line 4 "Disk Filename`t`t`t`t: " $DiskVersion.diskFileName
							Line 0 ""
						}
						If($HTML)
						{
							$rowdata = @()
							$columnHeaders = @("Version",($global:htmlsb),$BootFromVersion,$htmlwhite)
							If($DiskVersion.version -gt $Script:farm.maxVersions -and $VersionFlag -eq $False)
							{
								$rowdata += @(,('Version of vDisk is $($DiskVersion.version) which is greater than the limit of $($Script:farm.maxVersions)',($global:htmlsb),"Consider merging",$htmlwhite))
							}
							$rowdata += @(,('Created',($global:htmlsb),$DiskVersion.createDate,$htmlwhite))
							If(![String]::IsNullOrEmpty($DiskVersion.scheduledDate))
							{
								$rowdata += @(,('Released',($global:htmlsb),$DiskVersion.scheduledDate,$htmlwhite))
							}
							$rowdata += @(,('Devices',($global:htmlsb),$DiskVersion.deviceCount,$htmlwhite))
							$rowdata += @(,('Access',($global:htmlsb),$access,$htmlwhite))
							$rowdata += @(,('Type',($global:htmlsb),$DiskVersionType,$htmlwhite))
							If(![String]::IsNullOrEmpty($DiskVersion.description))
							{
								$rowdata += @(,('Properties',($global:htmlsb),$DiskVersion.description,$htmlwhite))
							}
							$rowdata += @(,('Can Delete',($global:htmlsb),$canDelete,$htmlwhite))
							$rowdata += @(,('Can Merge',($global:htmlsb),$canMerge,$htmlwhite))
							$rowdata += @(,('Can Merge Base',($global:htmlsb),$canMergeBase,$htmlwhite))
							$rowdata += @(,('Can Promote',($global:htmlsb),$canPromote,$htmlwhite))
							$rowdata += @(,('Can Revert back to Test',($global:htmlsb),$canRevertTest,$htmlwhite))
							$rowdata += @(,('Can Revert back to Maintenance',($global:htmlsb),$canRevertMaintenance,$htmlwhite))
							$rowdata += @(,('Can Set Scheduled Date',($global:htmlsb),$canSetScheduledDate,$htmlwhite))
							$rowdata += @(,('Can Override',($global:htmlsb),$canOverride,$htmlwhite))
							$rowdata += @(,('Is Pending',($global:htmlsb),$isPending,$htmlwhite))
							$rowdata += @(,('Replication Status',($global:htmlsb),$goodInventoryStatus,$htmlwhite))
							$rowdata += @(,('Disk Filename',($global:htmlsb),$DiskVersion.diskFileName,$htmlwhite))
					
							$msg = ""
							FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
							WriteHTMLLine 0 0 " "
						}
					}
				}
			}
			Else
			{
				If($MSWord -or $PDF)
				{
					WriteWordLine 0 0 "Disk Version information could not be retrieved"
					WriteWordLine 0 0 "Error returned is " $error[0].FullyQualifiedErrorId.Split(',')[0].Trim()
				}
				If($Text)
				{
					Line 0 "Disk Version information could not be retrieved"
					Line 0 "Error returned is " $error[0].FullyQualifiedErrorId.Split(',')[0].Trim()
				}
				If($HTML)
				{
					WriteHTMLLine 0 0 "Disk Version information could not be retrieved"
					WriteHTMLLine 0 0 "Error returned is " $error[0].FullyQualifiedErrorId.Split(',')[0].Trim()
				}
			}
			
			#process vDisk Load Balancing Menu
			Write-Verbose "$(Get-Date -Format G): `t`t`tProcessing vDisk Load Balancing Menu"
			If($Disk.rebalanceEnabled -eq "1")
			{
				$rebalanceEnabled = "Yes"
			}
			Else
			{
				$rebalanceEnabled = "No"
			}

			Switch ($Disk.subnetAffinity)
			{
				"0"		{$subnetAffinity = "None"; Break}
				"1"		{$subnetAffinity = "Best Effort"; Break}
				"2"		{$subnetAffinity = "Fixed"; Break}
				Default {$subnetAffinity = "Subnet Affinity could not be determined: $($Disk.subnetAffinity)"; Break}
			}

			If($MSWord -or $PDF)
			{
				WriteWordLine 3 0 "vDisk Load Balancing"
				[System.Collections.Hashtable[]] $ScriptInformation = @()
				If(![String]::IsNullOrEmpty($Disk.serverName))
				{
					$ScriptInformation += @{ Data = "Use this server to provide the vDisk"; Value = $Disk.serverName; }
				}
				Else
				{
					$ScriptInformation += @{ Data = "Subnet Affinity"; Value = $subnetAffinity; }
					$ScriptInformation += @{ Data = "Rebalance Enabled"; Value = $rebalanceEnabled; }
					If($Disk.rebalanceEnabled)
					{
						$ScriptInformation += @{ Data = "Trigger Percent"; Value = $Disk.rebalanceTriggerPercent; }
					}
				}

				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitContent;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
			}
			If($Text)
			{
				Line 3 "vDisk Load Balancing"
				If(![String]::IsNullOrEmpty($Disk.serverName))
				{
					Line 4 "Use this server to provide the vDisk: " $Disk.serverName
				}
				Else
				{
					Line 4 "Subnet Affinity`t`t: " $subnetAffinity
					Line 4 "Rebalance Enabled`t: " $rebalanceEnabled
					If($Disk.rebalanceEnabled)
					{
						Line 4 "Trigger Percent`t`t: $($Disk.rebalanceTriggerPercent)"
					}
				}
				Line 0 ""
			}
			If($HTML)
			{
				WriteHTMLLine 3 0 "vDisk Load Balancing"
				$rowdata = @()
				If(![String]::IsNullOrEmpty($Disk.serverName))
				{
					$columnHeaders = @("Use this server to provide the vDisk",($global:htmlsb),$Disk.serverName,$htmlwhite)
				}
				Else
				{
					$columnHeaders = @("Subnet Affinity",($global:htmlsb),$subnetAffinity,$htmlwhite)
					$rowdata += @(,('Rebalance Enabled',($global:htmlsb),$rebalanceEnabled,$htmlwhite))
					If($Disk.rebalanceEnabled)
					{
						$rowdata += @(,('Trigger Percent',($global:htmlsb),"$($Disk.rebalanceTriggerPercent)",$htmlwhite))
					}
				}
				
				$msg = ""
				FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
				WriteHTMLLine 0 0 ""
			}
		}

		# http://blogs.citrix.com/2013/07/03/pvs-internals-2-how-to-properly-size-your-memory/
		[decimal]$XDRecRAM = ((2 + ($NumberofvDisks * 2)) * 1.15)
		$XDRecRAM = "{0:N0}" -f $XDRecRAM

		[decimal]$XARecRAM = ((2 + ($NumberofvDisks * 4)) * 1.15)
		$XARecRAM = "{0:N0}" -f $XARecRAM

		[decimal]$XDXARecRAM = ((2 + (($NumberofvDisks * 4) + ($NumberofvDisks * 2))) * 1.15)
		$XDXARecRAM = "{0:N0}" -f $XDXARecRAM

		If($MSWord -or $PDF)
		{
			[System.Collections.Hashtable[]] $ScriptInformation = @()
			$ScriptInformation += @{ Data = "Number of vDisks that are Enabled and have active connections"; Value = $NumberofvDisks.ToString(); }
			$ScriptInformation += @{ Data = ""; Value = ""; }
			$ScriptInformation += @{ Data = "Recommended RAM for each PVS Server using XenDesktop vDisks"; Value = "$($XDRecRAM)GB"; }
			$ScriptInformation += @{ Data = "Recommended RAM for each PVS Server using XenApp vDisks"; Value = "$($XARecRAM)GB"; }
			$ScriptInformation += @{ Data = "Recommended RAM for each PVS Server using XA & XD vDisks"; Value = "$($XDXARecRAM)GB"; }

			$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data,Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitContent;

			SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 "This script is not able to tell if a vDisk is running XenDesktop or XenApp"
			WriteWordLine 0 0 "The RAM calculation is done based on both scenarios. The original formula is"
			WriteWordLine 0 0 "2GB + (#XA_vDisk * 4GB) + (#XD_vDisk * 2GB) + 15% (Buffer)"
			WriteWordLine 0 0 "PVS Internals 2 - How to properly size your memory by Martin Zugec"
			WriteWordLine 0 0 "https://www.citrix.com/blogs/2013/07/03/pvs-internals-2-how-to-properly-size-your-memory/"
			WriteWordLine 0 0 ""
		}
		If($Text)
		{
			Line 1 "Number of vDisks that are Enabled and have active connections: " $NumberofvDisks.ToString()
			Line 0 ""
			Line 1 "Recommended RAM for each PVS Server using XenDesktop vDisks  : $($XDRecRAM)GB"
			Line 1 "Recommended RAM for each PVS Server using XenApp vDisks      : $($XARecRAM)GB"
			Line 1 "Recommended RAM for each PVS Server using XA & XD vDisks     : $($XDXARecRAM)GB"
			Line 0 ""
			Line 1 "This script is not able to tell if a vDisk is running XenDesktop or XenApp."
			Line 1 "The RAM calculation is done based on both scenarios. The original formula is:"
			Line 1 "2GB + (#XA_vDisk * 4GB) + (#XD_vDisk * 2GB) + 15% (Buffer)"
			Line 1 'PVS Internals 2 - How to properly size your memory by Martin Zugec'
			Line 1 'https://www.citrix.com/blogs/2013/07/03/pvs-internals-2-how-to-properly-size-your-memory/'
			Line 0 ""
		}
		If($HTML)
		{
			$rowdata = @()
			$columnHeaders = @("Number of vDisks that are Enabled and have active connections",($global:htmlsb),$NumberofvDisks.ToString(),$htmlwhite)
			$rowdata += @(,('',($global:htmlsb),"",$htmlwhite))
			$rowdata += @(,('Recommended RAM for each PVS Server using XenDesktop vDisks',($global:htmlsb),"$($XDRecRAM)GB",$htmlwhite))
			$rowdata += @(,('Recommended RAM for each PVS Server using XenApp vDisks',($global:htmlsb),"$($XARecRAM)GB",$htmlwhite))
			$rowdata += @(,('Recommended RAM for each PVS Server using XA & XD vDisks',($global:htmlsb),"$($XDXARecRAM)GB",$htmlwhite))
			$msg = ""
			FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
			WriteHTMLLine 0 0 "This script is not able to tell if a vDisk is running XenDesktop or XenApp"
			WriteHTMLLine 0 0 "The RAM calculation is done based on both scenarios. The original formula is"
			WriteHTMLLine 0 0 "2GB + (#XA_vDisk * 4GB) + (#XD_vDisk * 2GB) + 15% (Buffer)"
			WriteHTMLLine 0 0 "PVS Internals 2 - How to properly size your memory by Martin Zugec"
			WriteHTMLLine 0 0 "https://www.citrix.com/blogs/2013/07/03/pvs-internals-2-how-to-properly-size-your-memory/"
			WriteHTMLLine 0 0 ""
		}
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 ""
			WriteWordLine 0 0 "***No vDisks were found in the Farm***" "" $Null 0 $False $True
			WriteWordLine 0 0 ""
		}
		If($Text)
		{
			Line 0 ""
			Line 1 "***No vDisks were found in the Farm***"
			Line 0 ""
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 ""
			WriteHTMLLine 0 0 "***No vDisks were found in the Farm***" -Option $htmlBold
			WriteHTMLLine 0 0 ""
		}
	}
}
#endregion

#region process stores functions
Function ProcessStores
{
	#process the stores now
	Write-Verbose "$(Get-Date -Format G): `tProcessing Stores"

	If($MSWord -or $PDF)
	{
		$selection.InsertNewPage()
		WriteWordLine 1 0 "Store Properties"
	}
	If($Text)
	{
		Line 0 "Store Properties"
	}
	If($HTML)
	{
		WriteHTMLLine 1 0 "Store Properties"
	}

	$GetWhat = "Store"
	$GetParam = ""
	$ErrorTxt = "Farm Store information"
	$Stores = BuildPVSObject $GetWhat $GetParam $ErrorTxt
	If($Null -ne $Stores)
	{
		ForEach($Store in $Stores)
		{
			Write-Verbose "$(Get-Date -Format G): `t`tProcessing Store $($Store.StoreName)"
			If($MSWord -or $PDF)
			{
				WriteWordLine 2 0 "Name: " $Store.StoreName
			}
			If($Text)
			{
				Line 1 "Name: " $Store.StoreName
			}
			If($HTML)
			{
				WriteHTMLLine 2 0 "Name: " $Store.StoreName
			}
			
			Write-Verbose "$(Get-Date -Format G): `t`t`tProcessing Servers Tab"
			If($MSWord -or $PDF)
			{
				WriteWordLine 3 0 "Servers"
			}
			If($Text)
			{
				Line 1 "Servers"
			}
			If($HTML)
			{
				WriteHTMLLine 3 0 "Servers"
			}

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
					Write-Verbose "$(Get-Date -Format G): `t`t`t`tProcessing Server $($Server.serverName)"
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

			If($MSWord -or $PDF)
			{
				WriteWordLine 4 0 "Servers that provide this store"
				[System.Collections.Hashtable[]] $ItemsWordTable = @();
			}
			If($Text)
			{
				Line 2 "Servers that provide this store:"
			}
			If($HTML)
			{
				$rowdata = @()
			}

			ForEach($StoreServer in $StoreServers)
			{
				If($MSWord -or $PDF)
				{
					If($MSWord -or $PDF)
					{
						$WordTableRowHash = @{ 
						StoreServer = $StoreServer;}
						$ItemsWordTable += $WordTableRowHash;
					}
				}
				If($Text)
				{
					Line 3 $StoreServer
				}
				If($HTML)
				{
					$rowdata += @(,(
					$StoreServer,$htmlwhite))
				}
			}

			If($MSWord -or $PDF)
			{
				If($ItemsWordTable.Count -gt 0)
				{
					$Table = AddWordTable -Hashtable $ItemsWordTable `
					-Columns StoreServer `
					-Headers "Store Server" `
					-AutoFit $wdAutoFitContent;

					SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

					$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

					FindWordDocumentEnd
					$Table = $Null
					$ItemsWordTable = $Null
					WriteWordLine 0 0 ""
				}
			}
			If($Text)
			{
				Line 0 ""
			}
			If($HTML)
			{
				If($Rowdata.Count -gt 0)
				{
					$columnHeaders = @(
						"Store Server",($global:htmlsb))
									
					$msg = "Servers that provide this store"
					FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
					WriteHTMLLine 0 0 " "
				}
			}

			Write-Verbose "$(Get-Date -Format G): `t`t`tProcessing Paths Tab"
			If($MSWord -or $PDF)
			{
				WriteWordLine 3 0 "Paths"
				[System.Collections.Hashtable[]] $ItemsWordTable = @();
			}
			If($Text)
			{
				Line 1 "Paths"
			}
			If($HTML)
			{
				WriteHTMLLine 3 0 "Paths"
				$rowdata = @()
			}

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
					If($MSWord -or $PDF)
					{
						$WordTableRowHash = @{ 
						StorePath = "Default store path: $($Store.path) on server $StoreServer";
						PathStatus = "Valid";}
						$ItemsWordTable += $WordTableRowHash;
					}
					If($Text)
					{
						Line 2 "Default store path: $($Store.path) on server $StoreServer is valid"
					}
					If($HTML)
					{
						$rowdata += @(,(
						"Default store path: $($Store.path) on server $StoreServer",$htmlwhite,
						"Valid",$htmlwhite))
					}
				}
				Else
				{
					If($MSWord -or $PDF)
					{
						$WordTableRowHash = @{ 
						StorePath = "Default store path: $($Store.path) on server $StoreServer";
						PathStatus = "Not Valid";}
						$ItemsWordTable += $WordTableRowHash;
					}
					If($Text)
					{
						Line 2 "Default store path: $($Store.path) on server $StoreServer is not valid"
					}
					If($HTML)
					{
						$rowdata += @(,(
						"Default store path: $($Store.path) on server $StoreServer",$htmlwhite,
						"Not Valid",$htmlwhite))
					}
				}
			}

			If($MSWord -or $PDF)
			{
				If($ItemsWordTable.Count -gt 0)
				{
					$Table = AddWordTable -Hashtable $ItemsWordTable `
					-Columns StorePath, PathStatus `
					-Headers "Store Path and Server", "Path Status" `
					-AutoFit $wdAutoFitContent;

					SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

					$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

					FindWordDocumentEnd
					$Table = $Null
					$ItemsWordTable = $Null
					WriteWordLine 0 0 ""
				}
			}
			If($Text)
			{
				Line 0 ""
			}
			If($HTML)
			{
				If($Rowdata.Count -gt 0)
				{
					$columnHeaders = @(
						"Store Path and Server",($global:htmlsb),
						"Path Status",($global:htmlsb))
									
					$msg = ""
					FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
					WriteHTMLLine 0 0 " "
				}
			}

			If(![String]::IsNullOrEmpty($Store.cachePath))
			{
				If($MSWord -or $PDF)
				{
					WriteWordLine 3 0 "Default write-cache paths"
					[System.Collections.Hashtable[]] $ItemsWordTable = @();
				}
				If($Text)
				{
					Line 2 "Default write-cache paths: "
				}
				If($HTML)
				{
					WriteHTMLLine 3 0 "Default write-cache paths"
					$rowdata = @()
				}

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
							If($MSWord -or $PDF)
							{
								$WordTableRowHash = @{ 
								StorePath = "Write Cache Path $($WCPath) on server $StoreServer";
								PathStatus = "Valid";}
								$ItemsWordTable += $WordTableRowHash;
							}
							If($Text)
							{
								Line 2 "Write Cache Path $($WCPath) on server $StoreServer is valid"
							}
							If($HTML)
							{
								$rowdata += @(,(
								"Write Cache Path $($WCPath) on server $StoreServer",$htmlwhite,
								"Valid",$htmlwhite))
							}
						}
						Else
						{
							If($MSWord -or $PDF)
							{
								$WordTableRowHash = @{ 
								StorePath = "Write Cache Path $($WCPath) on server $StoreServer";
								PathStatus = "Not Valid";}
								$ItemsWordTable += $WordTableRowHash;
							}
							If($Text)
							{
								Line 2 "Write Cache Path $($WCPath) on server $StoreServer is not valid"
							}
							If($HTML)
							{
								$rowdata += @(,(
								"Write Cache Path $($WCPath) on server $StoreServer",$htmlwhite,
								"Not Valid",$htmlwhite))
							}
						}
					}
				}
			}
			Else
			{
				If($MSWord -or $PDF)
				{
					[System.Collections.Hashtable[]] $ItemsWordTable = @();
					$WordTableRowHash = @{ 
					StorePath = "Using the default write-cache path of $($Store.Path)\WriteCache";
					PathStatus = "N/A";}
					$ItemsWordTable += $WordTableRowHash;
				}
				If($Text)
				{
					Line 2 "Using the default write-cache path of $($Store.Path)\WriteCache"
				}
				If($HTML)
				{
					$rowdata = @()
					$rowdata += @(,(
					"Using the default write-cache path of $($Store.Path)\WriteCache",$htmlwhite,
					"N/A",$htmlwhite))
				}
			}

			If($MSWord -or $PDF)
			{
				If($ItemsWordTable.Count -gt 0)
				{
					$Table = AddWordTable -Hashtable $ItemsWordTable `
					-Columns StorePath, PathStatus `
					-Headers "Write Cache Path and Server", "Path Status" `
					-AutoFit $wdAutoFitContent;

					SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

					$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

					FindWordDocumentEnd
					$Table = $Null
					$ItemsWordTable = $Null
					WriteWordLine 0 0 ""
				}
			}
			If($Text)
			{
				Line 0 ""
			}
			If($HTML)
			{
				If($Rowdata.Count -gt 0)
				{
					$columnHeaders = @(
						"Write Cache Path and Server",($global:htmlsb),
						"Path Status",($global:htmlsb))
									
					$msg = ""
					FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
					WriteHTMLLine 0 0 " "
				}
			}
		}
	}
	Else
	{
		$txt = "There are no Stores configured"
		OutputWarning $txt
	}

	If($MSWord -or $PDF)
	{
		WriteWordLine 0 0 ""
	}
	If($Text)
	{
		Line 0 ""
	}
	If($HTML)
	{
		WriteHTMLLine 0 0 ""
	}
	Write-Verbose "$(Get-Date -Format G): "
}
#endregion

#region appendix A
Function OutputAppendixA
{
	Write-Verbose "$(Get-Date -Format G): Create Appendix A Advanced Server Items (Server/Network)"
	#sort the array by servername
	$Script:AdvancedItems1 = $Script:AdvancedItems1 | Sort-Object ServerName
	
	If($CSV)
	{
		$File = "$($Script:pwdpath)\$($Script:farm.FarmName)_HealthCheckV2_AppendixA_AdvancedServerItems1.csv"
		$Script:AdvancedItems1 | Export-CSV -Force -Encoding ASCII -NoTypeInformation -Path $File
	}

	If($MSWord -or $PDF)
	{
		$selection.InsertNewPage()
		WriteWordLine 1 0 "Appendix A - Advanced Server Items (Server/Network)"
		[System.Collections.Hashtable[]] $ItemsWordTable = @();
	}
	If($Text)
	{
		Line 0 "Appendix A - Advanced Server Items (Server/Network)"
		Line 0 ""
		Line 1 "Server Name      Threads  Buffers  Server   Local       Remote      Ethernet  IO     Enable      "
		Line 1 "                 per      per      Cache    Concurrent  Concurrent  MTU       Burst  Non-blocking"
		Line 1 "                 Port     Thread   Timeout  IO Limit    IO Limit              Size   IO          "
		Line 1 "================================================================================================="
	}
	If($HTML)
	{
		WriteHTMLLine 1 0 "Appendix A - Advanced Server Items (Server/Network)"
		$rowdata = @()
	}

	ForEach($Item in $Script:AdvancedItems1)
	{
		If($MSWord -or $PDF)
		{
			$WordTableRowHash = @{ 
			ServerName              = $Item.serverName;
			ThreadsPerPort          = $Item.threadsPerPort;
			BuffersPerThread        = $Item.buffersPerThread;
			ServerCacheTimeout      = $Item.serverCacheTimeout;
			LocalConcurrentIOLimit  = $Item.localConcurrentIoLimit;
			RemoteConcurrentIOLimit = $Item.remoteConcurrentIoLimit;
			EthernetMTU             = $Item.maxTransmissionUnits;
			IOBurstSize             = $Item.ioBurstSize;
			EnableNonBlockingIO     = $Item.nonBlockingIoEnabled;}
			$ItemsWordTable += $WordTableRowHash;
		}
		If($Text)
		{
			Line 1 ( "{0,-16} {1,-8} {2,-8} {3,-8} {4,-11} {5,-11} {6,-9} {7,-6} {8,-8}" -f `
			$Item.serverName, $Item.threadsPerPort, $Item.buffersPerThread, $Item.serverCacheTimeout, `
			$Item.localConcurrentIoLimit, $Item.remoteConcurrentIoLimit, $Item.maxTransmissionUnits, $Item.ioBurstSize, `
			$Item.nonBlockingIoEnabled )
		}
		If($HTML)
		{
			$rowdata += @(,(
			$Item.serverName,$htmlwhite,
			$Item.threadsPerPort,$htmlwhite,
			$Item.buffersPerThread,$htmlwhite,
			$Item.serverCacheTimeout,$htmlwhite,
			$Item.localConcurrentIoLimit,$htmlwhite,
			$Item.remoteConcurrentIoLimit,$htmlwhite,
			$Item.maxTransmissionUnits,$htmlwhite,
			$Item.ioBurstSize,$htmlwhite,
			$Item.nonBlockingIoEnabled,$htmlwhite))
		}
	}

	If($MSWord -or $PDF)
	{
		$Table = AddWordTable -Hashtable $ItemsWordTable `
		-Columns ServerName, 
			ThreadsPerPort, 
			BuffersPerThread, 
			ServerCacheTimeout, 
			LocalConcurrentIOLimit, 
			RemoteConcurrentIOLimit, 
			EthernetMTU, 
			IOBurstSize, 
			EnableNonBlockingIO	`
		-Headers "Server Name", 
			"Threads per Port", 
			"Buffers per Thread", 
			"Server Cache Timeout", 
			"Local Concurrent IO Limit", 
			"Remote Concurrent IO Limit", 
			"Ethernet MTU", 
			"IO Burst Size", 
			"Enable Non-blocking IO" `
		-AutoFit $wdAutoFitContent;

		SetWordCellFormat -Collection $Table -Size 9 -BackgroundColor $wdColorWhite
		SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		$ItemsWordTable = $Null
	}
	If($Text)
	{
		Line 0 ""
	}
	If($HTML)
	{
		$columnHeaders = @(
			"Server Name", ($global:htmlsb),
			"Threads per Port", ($global:htmlsb),
			"Buffers per Thread", ($global:htmlsb),
			"Server Cache Timeout", ($global:htmlsb),
			"Local Concurrent IO Limit", ($global:htmlsb),
			"Remote Concurrent IO Limit", ($global:htmlsb),
			"Ethernet MTU", ($global:htmlsb),
			"IO Burst Size", ($global:htmlsb),
			"Enable Non-blocking IO",($global:htmlsb))
						
		$msg = ""
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -tablewidth "600"
		WriteHTMLLine 0 0 " "
	}

	Write-Verbose "$(Get-Date -Format G): Finished Creating Appendix A - Advanced Server Items (Server/Network)"
	Write-Verbose "$(Get-Date -Format G): "
}
#endregion

#region appendix B
Function OutputAppendixB
{
	Write-Verbose "$(Get-Date -Format G): Create Appendix B Advanced Server Items (Pacing/Device)"
	#sort the array by servername
	$Script:AdvancedItems2 = $Script:AdvancedItems2 | Sort-Object ServerName
	
	If($CSV)
	{
		$File = "$($Script:pwdpath)\$($Script:farm.FarmName)_HealthCheckV2_AppendixB_AdvancedServerItems2.csv"
		$Script:AdvancedItems2 | Export-CSV -Force -Encoding ASCII -NoTypeInformation -Path $File
	}

	If($MSWord -or $PDF)
	{
		$selection.InsertNewPage()
		WriteWordLine 1 0 "Appendix B - Advanced Server Items (Pacing/Device)"
		[System.Collections.Hashtable[]] $ItemsWordTable = @();
	}
	If($Text)
	{
		Line 0 "Appendix B - Advanced Server Items (Pacing/Device)"
		Line 0 ""
		Line 1 "Server Name      Boot     Maximum  Maximum  vDisk     License"
		Line 1 "                 Pause    Boot     Devices  Creation  Timeout"
		Line 1 "                 Seconds  Time     Booting  Pacing           "
		Line 1 "============================================================="
		###### "123451234512345  9999999  9999999  9999999  99999999  9999999
	}
	If($HTML)
	{
		WriteHTMLLine 1 0 "Appendix B - Advanced Server Items (Pacing/Device)"
		$rowdata = @()
	}

	ForEach($Item in $Script:AdvancedItems2)
	{
		If($MSWord -or $PDF)
		{
			$WordTableRowHash = @{ 
			ServerName            = $Item.serverName;
			BootPauseSeconds      = $Item.bootPauseSeconds;
			MaximumBootSeconds    = $Item.maxBootSeconds;
			MaximumDevicesBooting = $Item.maxBootDevicesAllowed;
			vDiskCreationPacing   = $Item.vDiskCreatePacing;
			LicenseTimeout        = $Item.licenseTimeout;}
			$ItemsWordTable += $WordTableRowHash;
		}
		If($Text)
		{
			Line 1 ( "{0,-16} {1,-8} {2,-8} {3,-8} {4,-9} {5,-8}" -f `
			$Item.serverName, $Item.bootPauseSeconds, $Item.maxBootSeconds, $Item.maxBootDevicesAllowed, `
			$Item.vDiskCreatePacing, $Item.licenseTimeout )
		}
		If($HTML)
		{
			$rowdata += @(,(
				$Item.serverName, $htmlwhite,
				$Item.bootPauseSeconds, $htmlwhite,
				$Item.maxBootSeconds, $htmlwhite,
				$Item.maxBootDevicesAllowed, $htmlwhite,
				$Item.vDiskCreatePacing, $htmlwhite,
				$Item.licenseTimeout,$htmlwhite))
		}
	}

	If($MSWord -or $PDF)
	{
		$Table = AddWordTable -Hashtable $ItemsWordTable `
		-Columns ServerName, 
			BootPauseSeconds, 
			MaximumBootSeconds, 
			MaximumDevicesBooting, 
			vDiskCreationPacing, 
			LicenseTimeout `
		-Headers "Server Name", 
			"Boot Pause Seconds", 
			"Maximum Boot Time", 
			"Maximum Devices Booting", 
			"vDisk Creation Pacing", 
			"License Timeout" `
		-AutoFit $wdAutoFitContent;

		SetWordCellFormat -Collection $Table -Size 9 -BackgroundColor $wdColorWhite
		SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		$ItemsWordTable = $Null
	}
	If($Text)
	{
		Line 0 ""
	}
	If($HTML)
	{
		$columnHeaders = @(
			"Server Name", ($global:htmlsb),
			"Boot Pause Seconds", ($global:htmlsb),
			"Maximum Boot Time", ($global:htmlsb),
			"Maximum Devices Booting", ($global:htmlsb),
			"vDisk Creation Pacing", ($global:htmlsb),
			"License Timeout", ($global:htmlsb))
						
		$msg = ""
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -tablewidth "600"
		WriteHTMLLine 0 0 " "
	}

	Write-Verbose "$(Get-Date -Format G): Finished Creating Appendix B - Advanced Server Items (Pacing/Device)"
	Write-Verbose "$(Get-Date -Format G): "
}
#endregion

#region appendix C
Function OutputAppendixC
{
	Write-Verbose "$(Get-Date -Format G): Create Appendix C Config Wizard Items"

	#sort the array by servername
	$Script:ConfigWizItems = $Script:ConfigWizItems | Sort-Object ServerName
	
	If($CSV)
	{
		$File = "$($Script:pwdpath)\$($Script:farm.FarmName)_HealthCheckV2_AppendixC_ConfigWizardItems.csv"
		$Script:ConfigWizItems | Export-CSV -Force -Encoding ASCII -NoTypeInformation -Path $File
	}
	
	If($MSWord -or $PDF)
	{
		$selection.InsertNewPage()
		WriteWordLine 1 0 "Appendix C - Configuration Wizard Settings"
		[System.Collections.Hashtable[]] $ItemsWordTable = @();
	}
	If($Text)
	{
		Line 0 "Appendix C - Configuration Wizard Settings"
		Line 0 ""
		Line 1 "Server Name      DHCP        PXE        TFTP    User                                               " 
		Line 1 "                 Services    Services   Option  Account                                            "
		Line 1 "================================================================================================"
	}
	If($HTML)
	{
		WriteHTMLLine 1 0 "Appendix C - Configuration Wizard Settings"
		$rowdata = @()
	}

	If($Script:ConfigWizItems)
	{
		ForEach($Item in $Script:ConfigWizItems)
		{
			If($MSWord -or $PDF)
			{
			$WordTableRowHash = @{ 
			ServerName   = $Item.serverName;
			DHCPServices = $Item.DHCPServicesValue.ToString();
			PXEService   = $Item.PXEServicesValue.ToString();
			TFTPOption   = $Item.TFTPOptionValue.ToString();
			UserAccount  = $Item.UserAccount;}
			$ItemsWordTable += $WordTableRowHash;
			}
			If($Text)
			{
				Line 1 ( "{0,-16} {1,-11} {2,-9} {3,-7} {4,-50}" -f `
				$Item.serverName, $Item.DHCPServicesValue.ToString(), $Item.PXEServicesValue.ToString(), $Item.TFTPOptionValue.ToString(), `
				$Item.UserAccount )
			}
			If($HTML)
			{
			$rowdata += @(,(
			$Item.serverName,$htmlwhite,
			$Item.DHCPServicesValue.ToString(),$htmlwhite,
			$Item.PXEServicesValue.ToString(),$htmlwhite,
			$Item.TFTPOptionValue.ToString(),$htmlwhite,
			$Item.UserAccount,$htmlwhite))
			}
		}

		If($MSWord -or $PDF)
		{
			$Table = AddWordTable -Hashtable $ItemsWordTable `
			-Columns ServerName, 
				DHCPServices, 
				PXEService, 
				TFTPOption, 
				UserAccount	`
			-Headers "Server Name", 
				"DHCP Services", 
				"PXE Services", 
				"TFTP Option", 
				"User Account" `
			-AutoFit $wdAutoFitContent;

			SetWordCellFormat -Collection $Table -Size 9 -BackgroundColor $wdColorWhite
			SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			$ItemsWordTable = $Null
		}
		If($Text)
		{
			Line 0 ""
		}
		If($HTML)
		{
			$columnHeaders = @(
				"Server Name", ($global:htmlsb),
				"DHCP Services", ($global:htmlsb),
				"PXE Services", ($global:htmlsb),
				"TFTP Option", ($global:htmlsb),
				"User Account",($global:htmlsb))
							
			$msg = ""
			FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -tablewidth "600"
			WriteHTMLLine 0 0 " "
		}
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 ""
			WriteWordLine 0 0 "***No Configuration Wizard Settings were found***" "" $Null 0 $False $True
			WriteWordLine 0 0 ""
		}
		If($Text)
		{
			Line 0 ""
			Line 1 "***No Configuration Wizard Settings were found***"
			Line 0 ""
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 ""
			WriteHTMLLine 0 0 "***No Configuration Wizard Settings were found***" -Option $htmlBold
			WriteHTMLLine 0 0 ""
		}
	}

	Write-Verbose "$(Get-Date -Format G): Finished Creating Appendix C - Config Wizard Items"
	Write-Verbose "$(Get-Date -Format G): "
}
#endregion

#region appendix D
Function OutputAppendixD
{
	Write-Verbose "$(Get-Date -Format G): Create Appendix D Server Bootstrap Items"

	#sort the array by bootstrapname and servername
	$Script:BootstrapItems = $Script:BootstrapItems | Sort-Object BootstrapName, ServerName
	
	If($CSV)
	{
		$File = "$($Script:pwdpath)\$($Script:farm.FarmName)_HealthCheckV2_AppendixD_ServerBootstrapItems.csv"
		$Script:BootstrapItems | Export-CSV -Force -Encoding ASCII -NoTypeInformation -Path $File
	}
	
	If($MSWord -or $PDF)
	{
		$selection.InsertNewPage()
		WriteWordLine 1 0 "Appendix D - Server Bootstrap Items"
		[System.Collections.Hashtable[]] $ItemsWordTable = @();
	}
	If($Text)
	{
		Line 0 "Appendix D - Server Bootstrap Items"
		Line 0 ""
		Line 1 "Bootstrap Name   Server Name      IP1              IP2              IP3              IP4" 
		Line 1 "===================================================================================================="
		########123456789012345  XXXXXXXXXXXXXXXX 123.123.123.123  123.123.123.123  123.123.123.123  123.123.123.123
	}
	If($HTML)
	{
		WriteHTMLLine 1 0 "Appendix D - Server Bootstrap Items"
		$rowdata = @()
	}

	If($Script:BootstrapItems)
	{
		ForEach($Item in $Script:BootstrapItems)
		{
			If($MSWord -or $PDF)
			{
				$WordTableRowHash = @{ 
				BootstrapName = $Item.BootstrapName;
				serverName    = $Item.serverName;
				IP1           = $Item.IP1;
				IP2           = $Item.IP2;
				IP3           = $Item.IP3;
				IP4           = $Item.IP4;}
				$ItemsWordTable += $WordTableRowHash;
			}
			If($Text)
			{
				Line 1 ( "{0,-16} {1,-16} {2,-16} {3,-16} {4,-16} {5,-16}" -f `
				$Item.BootstrapName, $Item.serverName, $Item.IP1, $Item.IP2, $Item.IP3, $Item.IP4 )
			}
			If($HTML)
			{
				$rowdata += @(,(
				$Item.BootstrapName,$htmlwhite,
				$Item.serverName,$htmlwhite,
				$Item.IP1,$htmlwhite,
				$Item.IP2,$htmlwhite,
				$Item.IP3,$htmlwhite,
				$Item.IP4,$htmlwhite))
			}
		}
	
		If($MSWord -or $PDF)
		{
			$Table = AddWordTable -Hashtable $ItemsWordTable `
			-Columns BootstrapName, 
				serverName, 
				IP1, 
				IP2, 
				IP3, 
				IP4	`
			-Headers "Bootstrap Name", 
				"Server Name", 
				"IP1", 
				"IP2", 
				"IP3", 
				"IP4" `
			-AutoFit $wdAutoFitContent;

			SetWordCellFormat -Collection $Table -Size 9 -BackgroundColor $wdColorWhite
			SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			$ItemsWordTable = $Null
		}
		If($Text)
		{
			Line 0 ""
		}
		If($HTML)
		{
			$columnHeaders = @(
				"Bootstrap Name", ($global:htmlsb),
				"Server Name", ($global:htmlsb),
				"IP1", ($global:htmlsb),
				"IP2", ($global:htmlsb),
				"IP3", ($global:htmlsb),
				"IP4",($global:htmlsb))
							
			$msg = ""
			FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -tablewidth "600"
			WriteHTMLLine 0 0 " "
		}
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 ""
			WriteWordLine 0 0 "***No Server Bootstrap Items were found***" "" $Null 0 $False $True
			WriteWordLine 0 0 ""
		}
		If($Text)
		{
			Line 0 ""
			Line 1 "***No Server Bootstrap Items were found***"
			Line 0 ""
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 ""
			WriteHTMLLine 0 0 "***No Server Bootstrap Items were found***" -Option $htmlBold
			WriteHTMLLine 0 0 ""
		}
	}

	Write-Verbose "$(Get-Date -Format G): Finished Creating Appendix D - Server Bootstrap Items"
	Write-Verbose "$(Get-Date -Format G): "
}
#endregion

#region appendix E
Function OutputAppendixE
{
	Write-Verbose "$(Get-Date -Format G): Create Appendix E DisableTaskOffload Setting"

	#sort the array by bootstrapname and servername
	$Script:TaskOffloadItems = $Script:TaskOffloadItems | Sort-Object ServerName
	
	If($CSV)
	{
		$File = "$($Script:pwdpath)\$($Script:farm.FarmName)_HealthCheckV2_AppendixE_DisableTaskOffloadSetting.csv"
		$Script:TaskOffloadItems | Export-CSV -Force -Encoding ASCII -NoTypeInformation -Path $File
	}
	
	If($MSWord -or $PDF)
	{
		$selection.InsertNewPage()
		WriteWordLine 1 0 "Appendix E - DisableTaskOffload Settings"
		[System.Collections.Hashtable[]] $ItemsWordTable = @();
		WriteWordLine 0 0 ""
		WriteWordLine 0 0 "Best Practices for Configuring Provisioning Services Server on a Network"
		WriteWordLine 0 0 "http://support.citrix.com/article/CTX117374"
		WriteWordLine 0 0 "This setting is not needed if you are running PVS 6.0 or later" "" $Null 0 $False $True	
		WriteWordLine 0 0 ""
	}
	If($Text)
	{
		Line 0 "Appendix E - DisableTaskOffload Settings"
		Line 0 ""
		Line 0 "Best Practices for Configuring Provisioning Services Server on a Network"
		Line 0 "http://support.citrix.com/article/CTX117374"
		Line 0 "This setting is not needed if you are running PVS 6.0 or later"
		Line 0 ""
		Line 1 "Server Name      DisableTaskOffload Setting" 
		Line 1 "==========================================="
	}
	If($HTML)
	{
		WriteHTMLLine 1 0 "Appendix E - DisableTaskOffload Settings"
		$rowdata = @()
		WriteHTMLLine 0 0 ""
		WriteHTMLLine 0 0 "Best Practices for Configuring Provisioning Services Server on a Network"
		WriteHTMLLine 0 0 "http://support.citrix.com/article/CTX117374"
		WriteHTMLLine 0 0 "This setting is not needed if you are running PVS 6.0 or later" "" $Null 2 $htmlbold
		WriteHTMLLine 0 0 ""
	}

	If($Script:TaskOffloadItems)
	{
		ForEach($Item in $Script:TaskOffloadItems)
		{
			If($MSWord -or $PDF)
			{
				$WordTableRowHash = @{ 
				ServerName       = $Item.serverName;
				TaskOffloadValue = $Item.TaskOffloadValue;}
				$ItemsWordTable += $WordTableRowHash;
			}
			If($Text)
			{
				Line 1 ( "{0,-16} {1,-16}" -f $Item.serverName, $Item.TaskOffloadValue )
			}
			If($HTML)
			{
				$rowdata += @(,(
				$Item.serverName,$htmlwhite,
				$Item.TaskOffloadValue,$htmlwhite))
			}
		}
	
		If($MSWord -or $PDF)
		{
			$Table = AddWordTable -Hashtable $ItemsWordTable `
			-Columns ServerName, 
				TaskOffloadValue `
			-Headers "Server Name", 
				"DisableTaskOffload Setting" `
			-AutoFit $wdAutoFitContent;

			SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			$ItemsWordTable = $Null
		}
		If($Text)
		{
			Line 0 ""
		}
		If($HTML)
		{
			$columnHeaders = @(
				"Server Name", ($global:htmlsb),
				"DisableTaskOffload Setting",($global:htmlsb))
							
			$msg = ""
			FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
			WriteHTMLLine 0 0 " "
		}
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 ""
			WriteWordLine 0 0 "***No DisableTaskOffload Settings were found***" "" $Null 0 $False $True
			WriteWordLine 0 0 ""
		}
		If($Text)
		{
			Line 0 ""
			Line 1 "***No DisableTaskOffload Settings were found***"
			Line 0 ""
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 ""
			WriteHTMLLine 0 0 "***No DisableTaskOffload Settings were found***" -Option $htmlBold
			WriteHTMLLine 0 0 ""
		}
	}

	Write-Verbose "$(Get-Date -Format G): Finished Creating Appendix E - DisableTaskOffload Setting"
	Write-Verbose "$(Get-Date -Format G): "
}
#endregion

#region appendix F
Function OutputAppendixF1
{
	Write-Verbose "$(Get-Date -Format G): Create Appendix F1 PVS Services"

	#sort the array by displayname and servername
	$Script:PVSServiceItems = $Script:PVSServiceItems | Sort-Object DisplayName, ServerName
	
	If($CSV)
	{
		#AppendixF1 and AppendixF2 items are contained in the same array
		$File = "$($Script:pwdpath)\$($Script:farm.FarmName)_HealthCheckV2_AppendixF_PVSServices.csv"
		$Script:PVSServiceItems | Export-CSV -Force -Encoding ASCII -NoTypeInformation -Path $File
	}
	
	If($MSWord -or $PDF)
	{
		$selection.InsertNewPage()
		WriteWordLine 1 0 "Appendix F1 - Server PVS Service Items"
		[System.Collections.Hashtable[]] $ItemsWordTable = @();
	}
	If($Text)
	{
		Line 0 "Appendix F1 - Server PVS Service Items"
		Line 0 ""
		Line 1 "Display Name                      Server Name      Service Name  Status Startup Type Started State   Log on as" 
		Line 1 "========================================================================================================================================"
		########123456789012345678901234567890123 123456789012345  1234567890123 123456 123456789012 1234567 
		#displayname, servername, name, status, startmode, started, startname, state 
	}
	If($HTML)
	{
		WriteHTMLLine 1 0 "Appendix F1 - Server PVS Service Items"
		$rowdata = @()
	}

	If($Script:PVSServiceItems)
	{
		ForEach($Item in $Script:PVSServiceItems)
		{
			If($MSWord -or $PDF)
			{
				$WordTableRowHash = @{ 
				DisplayName = $Item.DisplayName;
				ServerName  = $Item.serverName;
				ServiceName = $Item.ServiceName;
				Status      = $Item.Status;
				StartupType = $Item.StartMode;
				Started     = $Item.Started;
				State       = $Item.State;
				LogOnAs     = $Item.StartName;}
				$ItemsWordTable += $WordTableRowHash;
			}
			If($Text)
			{
				Line 1 ( "{0,-33} {1,-16} {2,-13} {3,-6} {4,-12} {5,-7} {6,-7} {7,-35}" -f `
				$Item.DisplayName, $Item.serverName, $Item.ServiceName, $Item.Status, $Item.StartMode, `
				$Item.Started, $Item.State, $Item.StartName )
			}
			If($HTML)
			{
				$rowdata += @(,(
				$Item.DisplayName,$htmlwhite,
				$Item.serverName,$htmlwhite,
				$Item.ServiceName,$htmlwhite,
				$Item.Status,$htmlwhite,
				$Item.StartMode,$htmlwhite,
				$Item.Started,$htmlwhite,
				$Item.State,$htmlwhite,
				$Item.StartName,$htmlwhite))
			}
		}
	
		If($MSWord -or $PDF)
		{
			$Table = AddWordTable -Hashtable $ItemsWordTable `
			-Columns DisplayName, 
				ServerName, 
				ServiceName, 
				Status, 
				StartupType, 
				Started, 
				State, 
				LogOnAs	`
			-Headers "Display Name", 
				"Server Name", 
				"Service Name", 
				"Status", 
				"Startup Type", 
				"Started", 
				"State", 
				"Log on as" `
			-AutoFit $wdAutoFitContent;

			SetWordCellFormat -Collection $Table -Size 9 -BackgroundColor $wdColorWhite
			SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			$ItemsWordTable = $Null
		}
		If($Text)
		{
			Line 0 ""
		}
		If($HTML)
		{
			$columnHeaders = @(
				"Display Name", ($global:htmlsb),
				"Server Name", ($global:htmlsb),
				"Service Name", ($global:htmlsb),
				"Status", ($global:htmlsb),
				"Startup Type", ($global:htmlsb),
				"Started", ($global:htmlsb),
				"State", ($global:htmlsb),
				"Log on as",($global:htmlsb))
							
			$msg = ""
			FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -tablewidth "600"
			WriteHTMLLine 0 0 " "
		}
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 ""
			WriteWordLine 0 0 "***No Server PVS Service Items were found***" "" $Null 0 $False $True
			WriteWordLine 0 0 ""
		}
		If($Text)
		{
			Line 0 ""
			Line 1 "***No Server PVS Service Items were found***"
			Line 0 ""
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 ""
			WriteHTMLLine 0 0 "***No Server PVS Service Items were found***" -Option $htmlBold
			WriteHTMLLine 0 0 ""
		}
	}

	Write-Verbose "$(Get-Date -Format G): Finished Creating Appendix F1 - PVS Services"
	Write-Verbose "$(Get-Date -Format G): "
}
#endregion

#region appendix F2
Function OutputAppendixF2
{
	Write-Verbose "$(Get-Date -Format G): Create Appendix F2 PVS Services Failure Actions"
	#array is already sorted in Function OutputAppendixF1
	
	If($MSWord -or $PDF)
	{
		$selection.InsertNewPage()
		WriteWordLine 1 0 "Appendix F2 - Server PVS Service Items Failure Actions"
		[System.Collections.Hashtable[]] $ItemsWordTable = @();
	}
	If($Text)
	{
		Line 0 "Appendix F2 - Server PVS Service Items Failure Actions"
		Line 0 ""
		Line 1 "Display Name                      Server Name      Service Name  Failure Action 1     Failure Action 2     Failure Action 3    " 
		Line 1 "==============================================================================================================================="
	}
	If($HTML)
	{
		WriteHTMLLine 1 0 "Appendix F2 - Server PVS Service Items Failure Actions"
		$rowdata = @()
	}

	If($Script:PVSServiceItems)
	{
		ForEach($Item in $Script:PVSServiceItems)
		{
			If($MSWord -or $PDF)
			{
				$WordTableRowHash = @{ 
				DisplayName    = $Item.DisplayName;
				serverName     = $Item.serverName;
				ServiceName    = $Item.ServiceName;
				FailureAction1 = $Item.FailureAction1;
				FailureAction2 = $Item.FailureAction2;
				FailureAction3 = $Item.FailureAction3;}
				$ItemsWordTable += $WordTableRowHash;
			}
			If($Text)
			{
				Line 1 ( "{0,-33} {1,-16} {2,-13} {3,-20} {4,-20} {5,-20}" -f `
				$Item.DisplayName, $Item.serverName, $Item.ServiceName, $Item.FailureAction1, $Item.FailureAction2, $Item.FailureAction3 )
			}
			If($HTML)
			{
				$rowdata += @(,(
				$Item.DisplayName,$htmlwhite,
				$Item.serverName,$htmlwhite,
				$Item.ServiceName,$htmlwhite,
				$Item.FailureAction1,$htmlwhite,
				$Item.FailureAction2,$htmlwhite,
				$Item.FailureAction3,$htmlwhite))
			}
		}
	
		If($MSWord -or $PDF)
		{
			$Table = AddWordTable -Hashtable $ItemsWordTable `
			-Columns DisplayName, 
				serverName, 
				ServiceName, 
				FailureAction1, 
				FailureAction2, 
				FailureAction3 `
			-Headers "Display Name", 
				"Server Name", 
				"Service Name", 
				"Failure Action 1", 
				"Failure Action 2", 
				"Failure Action 3" `
			-AutoFit $wdAutoFitContent;

			SetWordCellFormat -Collection $Table -Size 9 -BackgroundColor $wdColorWhite
			SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			$ItemsWordTable = $Null
		}
		If($Text)
		{
			Line 0 ""
		}
		If($HTML)
		{
			$columnHeaders = @(
				"Display Name", ($global:htmlsb),
				"Server Name", ($global:htmlsb),
				"Service Name", ($global:htmlsb),
				"Failure Action 1", ($global:htmlsb),
				"Failure Action 2", ($global:htmlsb),
				"Failure Action 3",($global:htmlsb))
							
			$msg = ""
			FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -tablewidth "600"
			WriteHTMLLine 0 0 " "
		}
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 ""
			WriteWordLine 0 0 "***No Server PVS Service Items Failure Actions were found***" "" $Null 0 $False $True
			WriteWordLine 0 0 ""
		}
		If($Text)
		{
			Line 0 ""
			Line 1 "***No Server PVS Service Items Failure Actions were found***"
			Line 0 ""
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 ""
			WriteHTMLLine 0 0 "***No Server PVS Service Items Failure Actions were found***" -Option $htmlBold
			WriteHTMLLine 0 0 ""
		}
	}

	Write-Verbose "$(Get-Date -Format G): Finished Creating Appendix F2 - PVS Services Failure Actions"
	Write-Verbose "$(Get-Date -Format G): "
}
#endregion

#region appendix G
Function OutputAppendixG
{
	Write-Verbose "$(Get-Date -Format G): Create Appendix G vDisks to Merge"

	#sort the array
	$Script:VersionsToMerge = $Script:VersionsToMerge | Sort-Object
	
	If($CSV)
	{
		$File = "$($Script:pwdpath)\$($Script:farm.FarmName)_HealthCheckV2_AppendixG_vDiskstoMerge.csv"
		$Script:VersionsToMerge | Export-CSV -Force -Encoding ASCII -NoTypeInformation -Path $File
	}
	
	If($MSWord -or $PDF)
	{
		$selection.InsertNewPage()
		WriteWordLine 1 0 "Appendix G - vDisks to Consider Merging"
		[System.Collections.Hashtable[]] $ItemsWordTable = @();
	}
	If($Text)
	{
		Line 0 "Appendix G - vDisks to Consider Merging"
		Line 0 ""
		Line 1 "vDisk Name" 
		Line 1 "========================================"
	}
	If($HTML)
	{
		WriteHTMLLine 1 0 "Appendix G - vDisks to Consider Merging"
		$rowdata = @()
	}

	If($Script:VersionsToMerge)
	{
		ForEach($Item in $Script:VersionsToMerge)
		{
			If($MSWord -or $PDF)
			{
				$WordTableRowHash = @{ 
				vDiskName = $Item.vDiskName;}
				$ItemsWordTable += $WordTableRowHash;
			}
			If($Text)
			{
				Line 1 ( "{0,-40}" -f $Item.vDiskName )
			}
			If($HTML)
			{
				$rowdata += @(,(
				$Item.vDiskName,$htmlwhite))
			}
		}

		If($MSWord -or $PDF)
		{
			$Table = AddWordTable -Hashtable $ItemsWordTable `
			-Columns vDiskName `
			-Headers "vDisk Name" `
			-AutoFit $wdAutoFitContent;

			SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			$ItemsWordTable = $Null
		}
		If($Text)
		{
			Line 0 ""
		}
		If($HTML)
		{
			$columnHeaders = @(
				"vDisk Name",($global:htmlsb))
							
			$msg = ""
			FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
			WriteHTMLLine 0 0 " "
		}
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 ""
			WriteWordLine 0 0 "***No vDisks to Consider Merging were found***" "" $Null 0 $False $True
			WriteWordLine 0 0 ""
		}
		If($Text)
		{
			Line 0 ""
			Line 1 "***No vDisks to Consider Merging were found***"
			Line 0 ""
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 ""
			WriteHTMLLine 0 0 "***No vDisks to Consider Mergings were found***" -Option $htmlBold
			WriteHTMLLine 0 0 ""
		}
	}

	Write-Verbose "$(Get-Date -Format G): Finished Creating Appendix G - vDisks to Merge"
	Write-Verbose "$(Get-Date -Format G): "
}
#endregion

#region appendix H
Function OutputAppendixH
{
	Write-Verbose "$(Get-Date -Format G): Create Appendix H Empty Device Collections"

	#sort the array
	$Script:EmptyDeviceCollections = $Script:EmptyDeviceCollections | Sort-Object CollectionName
	
	If($CSV)
	{
		$File = "$($Script:pwdpath)\$($Script:farm.FarmName)_HealthCheckV2_AppendixH_EmptyDeviceCollections.csv"
		$Script:EmptyDeviceCollections | Export-CSV -Force -Encoding ASCII -NoTypeInformation -Path $File
	}
	
	If($MSWord -or $PDF)
	{
		$selection.InsertNewPage()
		WriteWordLine 1 0 "Appendix H - Empty Device Collections"
		[System.Collections.Hashtable[]] $ItemsWordTable = @();
	}
	If($Text)
	{
		Line 0 "Appendix H - Empty Device Collections"
		Line 0 ""
		Line 1 "Device Collection Name" 
		Line 1 "=================================================="
	}
	If($HTML)
	{
		WriteHTMLLine 1 0 "Appendix H - Empty Device Collections"
		$rowdata = @()
	}

	If($Script:EmptyDeviceCollections)
	{
		ForEach($Item in $Script:EmptyDeviceCollections)
		{
			If($MSWord -or $PDF)
			{
				$WordTableRowHash = @{ 
				CollectionName = $Item.CollectionName;}
				$ItemsWordTable += $WordTableRowHash;
			}
			If($Text)
			{
				Line 1 ( "{0,-50}" -f $Item.CollectionName )
			}
			If($HTML)
			{
				$rowdata += @(,(
				$Item.CollectionName,$htmlwhite))
			}
		}
	
		If($MSWord -or $PDF)
		{
			$Table = AddWordTable -Hashtable $ItemsWordTable `
			-Columns CollectionName `
			-Headers "Device Collection Name" `
			-AutoFit $wdAutoFitContent;

			SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			$ItemsWordTable = $Null
		}
		If($Text)
		{
			Line 0 ""
		}
		If($HTML)
		{
			$columnHeaders = @(
				"Device Collection Name",($global:htmlsb))
							
			$msg = ""
			FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
			WriteHTMLLine 0 0 " "
		}
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 ""
			WriteWordLine 0 0 "***No Empty Device Collections were found***" "" $Null 0 $False $True
			WriteWordLine 0 0 ""
		}
		If($Text)
		{
			Line 0 ""
			Line 1 "***No Empty Device Collections were found***"
			Line 0 ""
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 ""
			WriteHTMLLine 0 0 "***No Empty Device Collections were found***" -Option $htmlBold
			WriteHTMLLine 0 0 ""
		}
	}

	Write-Verbose "$(Get-Date -Format G): Finished Creating Appendix G - Empty Device Collections"
	Write-Verbose "$(Get-Date -Format G): "
}
#endregion

#region appendix I 
Function ProcessvDisksWithNoAssociation
{
	Write-Verbose "$(Get-Date -Format G): Finding vDisks with no Target Device Associations"
	$UnassociatedvDisks = New-Object System.Collections.ArrayList
	$GetWhat = "diskLocator"
	$GetParam = ""
	$ErrorTxt = "Disk Locator information"
	$DiskLocators = BuildPVSObject $GetWhat $GetParam $ErrorTxt
	
	If($Null -eq $DiskLocators)
	{
		Write-Verbose "$(Get-Date -Format G): No DiskLocators Found"
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
			#Write-Verbose "$(Get-Date -Format G): Found $($UnassociatedvDisks.Count) vDisks with no Target Device Associations"
			OutputAppendixI $UnassociatedvDisks
		}
		Else
		{
			#Write-Verbose "$(Get-Date -Format G): All vDisks have Target Device Associations"
			Write-Verbose "$(Get-Date -Format G): "
			OutputAppendixI $Null
		}
	}
}

Function OutputAppendixI
{
	Param([array]$vDisks)

	Write-Verbose "$(Get-Date -Format G): Create Appendix I Unassociated vDisks"

	If($MSWord -or $PDF)
	{
		$selection.InsertNewPage()
		WriteWordLine 1 0 "Appendix I - vDisks with no Target Device Associations"
		[System.Collections.Hashtable[]] $ItemsWordTable = @();
	}
	If($Text)
	{
		Line 0 "Appendix I - vDisks with no Target Device Associations"
		Line 0 ""
		Line 1 "vDisk Name" 
		Line 1 "========================================"
	}
	If($HTML)
	{
		WriteHTMLLine 1 0 "Appendix I - vDisks with no Target Device Associations"
		$rowdata = @()
	}
	
	If($vDisks)
	{
		#sort the array
		$vDisks = $vDisks | Sort-Object
	
		If($CSV)
		{
			$File = "$($Script:pwdpath)\$($Script:farm.FarmName)_HealthCheckV2_AppendixI_UnassociatedvDisks.csv"
			$vDisks | Export-CSV -Force -Encoding ASCII -NoTypeInformation -Path $File
		}
	
		ForEach($Item in $vDisks)
		{
			If($MSWord -or $PDF)
			{
				$WordTableRowHash = @{ 
				vDiskName = $Item.vDiskName;}
				$ItemsWordTable += $WordTableRowHash;
			}
			If($Text)
			{
				Line 1 ( "{0,-40}" -f $Item.vDiskName )
			}
			If($HTML)
			{
				$rowdata += @(,(
				$Item.vDiskName,$htmlwhite))
			}
		}
	
		If($MSWord -or $PDF)
		{
			$Table = AddWordTable -Hashtable $ItemsWordTable `
			-Columns vDiskName `
			-Headers "vDisk Name" `
			-AutoFit $wdAutoFitContent;

			SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			$ItemsWordTable = $Null
		}
		If($Text)
		{
			Line 0 ""
		}
		If($HTML)
		{
			$columnHeaders = @(
				"vDisk Name",($global:htmlsb))
							
			$msg = ""
			FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
			WriteHTMLLine 0 0 " "
		}
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 ""
			WriteWordLine 0 0 "***No vDisks with no Target Device Associations were found***" "" $Null 0 $False $True
			WriteWordLine 0 0 ""
		}
		If($Text)
		{
			Line 0 ""
			Line 1 "***No vDisks with no Target Device Associations were found***"
			Line 0 ""
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 ""
			WriteHTMLLine 0 0 "***No vDisks with no Target Device Associations were found***" -Option $htmlBold
			WriteHTMLLine 0 0 ""
		}
	}

	Write-Verbose "$(Get-Date -Format G): Finished Creating Appendix I - Unassociated vDisks"
	Write-Verbose "$(Get-Date -Format G): "
}
#endregion

#region appendix J
Function OutputAppendixJ
{
	Write-Verbose "$(Get-Date -Format G): Create Appendix J Bad Streaming IP Addresses"

	#sort the array by bootstrapname and servername
	$Script:BadIPs = $Script:BadIPs | Sort-Object ServerName, IPAddress
	
	If($CSV)
	{
		$File = "$($Script:pwdpath)\$($Script:farm.FarmName)_HealthCheckV2_AppendixJ_BadStreamingIPAddresses.csv"
		$Script:BadIPs | Export-CSV -Force -Encoding ASCII -NoTypeInformation -Path $File
	}
	
	If($MSWord -or $PDF)
	{
		$selection.InsertNewPage()
		WriteWordLine 1 0 "Appendix J - Bad Streaming IP Addresses"
		WriteWordLine 0 0 "Streaming IP addresses that do not exist on the server"
		[System.Collections.Hashtable[]] $ItemsWordTable = @();
	}
	If($Text)
	{
		Line 0 "Appendix J - Bad Streaming IP Addresses"
		Line 0 "Streaming IP addresses that do not exist on the server"
	}
	If($HTML)
	{
		WriteHTMLLine 1 0 "Appendix J - Bad Streaming IP Addresses"
		WriteHTMLLine 0 0 "Streaming IP addresses that do not exist on the server"
		$rowdata = @()
	}

	If($Script:PVSVersion -eq "7")
	{
		If($Text)
		{
			Line 0 ""
			Line 1 "Server Name      Streaming IP Address" 
			Line 1 "====================================="
		}

		If($Script:BadIPs) 
		{
			ForEach($Item in $Script:BadIPs)
			{
				If($MSWord -or $PDF)
				{
					$WordTableRowHash = @{ 
					ServerName = $Item.serverName;
					IPAddress  = $Item.IPAddress;}
					$ItemsWordTable += $WordTableRowHash;
				}
				If($Text)
				{
					Line 1 ( "{0,-16} {1,-16}" -f $Item.serverName, $Item.IPAddress )
				}
				If($HTML)
				{
					$rowdata += @(,(
					$Item.serverName,$htmlwhite,
					$Item.IPAddress,$htmlwhite))
				}
			}
	
			If($MSWord -or $PDF)
			{
				$Table = AddWordTable -Hashtable $ItemsWordTable `
				-Columns ServerName, 
					IPAddress `
				-Headers "Server Name", 
					"Streaming IP Address" `
				-AutoFit $wdAutoFitContent;

				SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				$ItemsWordTable = $Null
			}
			If($Text)
			{
				Line 0 ""
			}
			If($HTML)
			{
				$columnHeaders = @(
					"Server Name", ($global:htmlsb),
					"Streaming IP Address",($global:htmlsb))
								
				$msg = ""
				FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
				WriteHTMLLine 0 0 " "
			}
		}
		Else
		{
			If($MSWord -or $PDF)
			{
				WriteWordLine 0 0 ""
				WriteWordLine 0 0 "***No Bad Streaming IP Addresses were found***" "" $Null 0 $False $True
				WriteWordLine 0 0 ""
			}
			If($Text)
			{
				Line 0 ""
				Line 1 "***No Bad Streaming IP Addresses were found***"
				Line 0 ""
			}
			If($HTML)
			{
				WriteHTMLLine 0 0 ""
				WriteHTMLLine 0 0 "***No Bad Streaming IP Addresses were found***" -Option $htmlBold
				WriteHTMLLine 0 0 ""
			}
		}
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			Line 1 "Unable to determine Bad Streaming IP Addresses for PVS versions earlier than 7.0"
		}
		If($Text)
		{
			Line 1 "Unable to determine Bad Streaming IP Addresses for PVS versions earlier than 7.0"
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 1 "Unable to determine Bad Streaming IP Addresses for PVS versions earlier than 7.0"
		}
	}

	Write-Verbose "$(Get-Date -Format G): Finished Creating Appendix J Bad Streaming IP Addresses"
	Write-Verbose "$(Get-Date -Format G): "
}
#endregion

#region appendix K
Function OutputAppendixK
{
	Write-Verbose "$(Get-Date -Format G): Create Appendix K Misc Registry Items"

	#sort the array by regkey, regvalue and servername
	$Script:MiscRegistryItems = $Script:MiscRegistryItems | Sort-Object RegKey, RegValue, ServerName
	
	If($CSV)
	{
		$File = "$($Script:pwdpath)\$($Script:farm.FarmName)_HealthCheckV2_AppendixK_MiscRegistryItems.csv"
		$Script:MiscRegistryItems | Export-CSV -Force -Encoding ASCII -NoTypeInformation -Path $File
	}
	
	If($MSWord -or $PDF)
	{
		$selection.InsertNewPage()
		WriteWordLine 1 0 "Appendix K - Misc Registry Items"
		WriteWordLine 0 0 "Miscellaneous Registry Items That May or May Not Exist on Servers"
		WriteWordLine 0 0 "These items may or may not be needed"
		WriteWordLine 0 0 "This Appendix is strictly for server comparison only"
		WriteWordLine 0 0 ""
		[System.Collections.Hashtable[]] $ItemsWordTable = @();
	}
	If($Text)
	{
		Line 0 "Appendix K - Misc Registry Items"
		Line 0 "Miscellaneous Registry Items That May or May Not Exist on Servers"
		Line 0 "These items may or may not be needed"
		Line 0 "This Appendix is strictly for server comparison only"
		Line 0 ""
		Line 1 "Registry Key                                                                                    Registry Value                                     Data                                                                                       Server Name    " 
		Line 1 "============================================================================================================================================================================================================================================================="
		#       12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345S12345678901234567890123456789012345678901234567890S123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890S123456789012345
	}
	If($HTML)
	{
		WriteHTMLLine 1 0 "Appendix K - Misc Registry Items"
		WriteHTMLLine 0 0 "Miscellaneous Registry Items That May or May Not Exist on Servers"
		WriteHTMLLine 0 0 "These items may or may not be needed"
		WriteHTMLLine 0 0 "This Appendix is strictly for server comparison only"
		WriteHTMLLine 0 0 ""
		$rowdata = @()
	}
	
	$Save = ""
	$First = $True
	If($Script:MiscRegistryItems)
	{
		ForEach($Item in $Script:MiscRegistryItems)
		{
			If(!$First -and $Save -ne "$($Item.RegKey.ToString())$($Item.RegValue.ToString())")
			{
				If($MSWord -or $PDF)
				{
					$WordTableRowHash = @{ 
					RegKey     = "";
					RegValue   = "";
					Value      = "";
					serverName = "";}
					$ItemsWordTable += $WordTableRowHash;
				}
				If($Text)
				{
					Line 0 ""
				}
				If($HTML)
				{
					$rowdata += @(,(
					"",$htmlwhite,
					"",$htmlwhite,
					"",$htmlwhite,
					"",$htmlwhite))
				}
			}

			If($MSWord -or $PDF)
			{
				$WordTableRowHash = @{ 
				RegKey     = $Item.RegKey;
				RegValue   = $Item.RegValue;
				Value      = $Item.Value;
				serverName = $Item.serverName;}
				$ItemsWordTable += $WordTableRowHash;
			}
			If($Text)
			{
				Line 1 ( "{0,-95} {1,-50} {2,-90} {3,-15}" -f `
				$Item.RegKey, $Item.RegValue, $Item.Value, $Item.serverName )
			}
			If($HTML)
			{
				$rowdata += @(,(
				$Item.RegKey,$htmlwhite,
				$Item.RegValue,$htmlwhite,
				$Item.Value,$htmlwhite,
				$Item.serverName,$htmlwhite))
			}

			$Save = "$($Item.RegKey.ToString())$($Item.RegValue.ToString())"
			If($First)
			{
				$First = $False
			}
		}
	
		If($MSWord -or $PDF)
		{
			$Table = AddWordTable -Hashtable $ItemsWordTable `
			-Columns RegKey, 
				RegValue, 
				Value, 
				serverName	`
			-Headers "Registry Key", 
				"Registry Value", 
				"Data", 
				"Server Name" `
			-AutoFit $wdAutoFitContent;

			SetWordCellFormat -Collection $Table -Size 9 -BackgroundColor $wdColorWhite
			SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			$ItemsWordTable = $Null
		}
		If($Text)
		{
			Line 0 ""
		}
		If($HTML)
		{
			$columnHeaders = @(
				"Registry Key", ($global:htmlsb),
				"Registry Value", ($global:htmlsb),
				"Data", ($global:htmlsb),
				"Server Name",($global:htmlsb))
							
			$msg = ""
			FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -tablewidth "600"
			WriteHTMLLine 0 0 " "
		}
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 ""
			WriteWordLine 0 0 "***No Misc Registry Items were found***" "" $Null 0 $False $True
			WriteWordLine 0 0 ""
		}
		If($Text)
		{
			Line 0 ""
			Line 1 "***No Misc Registry Items were found***"
			Line 0 ""
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 ""
			WriteHTMLLine 0 0 "***No Misc Registry Items were found***" -Option $htmlBold
			WriteHTMLLine 0 0 ""
		}
	}

	Write-Verbose "$(Get-Date -Format G): Finished Creating Appendix K Misc Registry Items"
	Write-Verbose "$(Get-Date -Format G): "
}
#endregion

#region appendix L
Function OutputAppendixL
{
	Write-Verbose "$(Get-Date -Format G): Create Appendix L vDisks Configured for Server-Side Caching"
	#sort the array 
	$Script:CacheOnServer = $Script:CacheOnServer | Sort-Object StoreName,SiteName,vDiskName
	
	If($CSV)
	{
		$File = "$($Script:pwdpath)\$($Script:farm.FarmName)_HealthCheckV2_AppendixL_vDisksConfiguredforServerSideCaching.csv"
		$Script:CacheOnServer | Export-CSV -Force -Encoding ASCII -NoTypeInformation -Path $File
	}

	If($MSWord -or $PDF)
	{
		$selection.InsertNewPage()
		WriteWordLine 1 0 "Appendix L - vDisks Configured for Server Side-Caching"
		[System.Collections.Hashtable[]] $ItemsWordTable = @();
	}
	If($Text)
	{
		Line 0 "Appendix L - vDisks Configured for Server Side-Caching"
		Line 0 ""
	}
	If($HTML)
	{
		WriteHTMLLine 1 0 "Appendix L - vDisks Configured for Server Side-Caching"
		$rowdata = @()
	}

	If($Script:CacheOnServer)
	{
		If($Text)
		{
			Line 1 "Store Name                Site Name                 vDisk Name               "
			Line 1 "============================================================================="
				   #1234567890123456789012345 1234567890123456789012345 1234567890123456789012345
		}

		ForEach($Item in $Script:CacheOnServer)
		{
			If($MSWord -or $PDF)
			{
				$WordTableRowHash = @{ 
				StoreName = $Item.StoreName;
				SiteName  = $Item.SiteName;
				vDiskName = $Item.vDiskName;}
				$ItemsWordTable += $WordTableRowHash;
			}
			If($Text)
			{
				Line 1 ( "{0,-25} {1,-25} {2,-25}" -f `
				$Item.StoreName, $Item.SiteName, $Item.vDiskName )
			}
			If($HTML)
			{
				$rowdata += @(,(
				$Item.StoreName,$htmlwhite,
				$Item.SiteName,$htmlwhite,
				$Item.vDiskName,$htmlwhite))
			}
		}
	
		If($MSWord -or $PDF)
		{
			$Table = AddWordTable -Hashtable $ItemsWordTable `
			-Columns StoreName, 
				SiteName, 
				vDiskName `
			-Headers "Store Name", 
				"Site Name", 
				"vDisk Name" `
			-AutoFit $wdAutoFitContent;

			SetWordCellFormat -Collection $Table -Size 9 -BackgroundColor $wdColorWhite
			SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			$ItemsWordTable = $Null
		}
		If($Text)
		{
			Line 0 ""
		}
		If($HTML)
		{
			$columnHeaders = @(
				"Store Name", ($global:htmlsb),
				"Site Name", ($global:htmlsb),
				"vDisk Name",($global:htmlsb))
							
			$msg = ""
			FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -tablewidth "600"
			WriteHTMLLine 0 0 " "
		}
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 ""
			WriteWordLine 0 0 "***No vDisks Configured for Server Side-Caching were found***" "" $Null 0 $False $True
			WriteWordLine 0 0 ""
		}
		If($Text)
		{
			Line 1 "***No vDisks Configured for Server Side-Caching were found***"
			Line 0 ""
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 ""
			WriteHTMLLine 0 0 "***No vDisks Configured for Server Side-Caching were found***" -Option $htmlBold
			WriteHTMLLine 0 0 ""
		}
	}
	
	Write-Verbose "$(Get-Date -Format G): Finished Creating Appendix L vDisks Configured for Server-Side Caching"
	Write-Verbose "$(Get-Date -Format G): "
}
#endregion

#region appendix M
Function OutputAppendixM
{
	#added in V1.16
	Write-Verbose "$(Get-Date -Format G): Create Appendix M Microsoft Hotfixes and Updates"

	#sort the array by hotfixid and servername
	$Script:MSHotfixes = $Script:MSHotfixes | Sort-Object HotFixID, ServerName
	
	If($CSV)
	{
		$File = "$($Script:pwdpath)\$($Script:farm.FarmName)_HealthCheckV2_AppendixM_MicrosoftHotfixesandUpdates.csv"
		$Script:MSHotfixes | Export-CSV -Force -Encoding ASCII -NoTypeInformation -Path $File
	}
	
	If($MSWord -or $PDF)
	{
		$selection.InsertNewPage()
		WriteWordLine 1 0 "Appendix M - Microsoft Hotfixes and Updates"
		[System.Collections.Hashtable[]] $ItemsWordTable = @();
	}
	If($Text)
	{
		Line 0 "Appendix M - Microsoft Hotfixes and Updates"
		Line 0 ""
		Line 1 "Hotfix ID                 Server Name     Caption                                       Description          Installed By                        Installed On Date     "
		Line 1 "======================================================================================================================================================================="
		#       1234567890123456789012345S123456789012345S123456789012345678901234567890123456789012345S12345678901234567890S12345678901234567890123456789012345S1234567890123456789012
		#                                                 http://support.microsoft.com/?kbid=2727528    Security Update      XXX-XX-XDDC01\xxxx.xxxxxx           00/00/0000 00:00:00 PM
		#		25                        15              45                                            20                   35                                  22
	}
	If($HTML)
	{
		WriteHTMLLine 1 0 "Appendix M - Microsoft Hotfixes and Updates"
		$rowdata = @()
	}
	
	$Save = ""
	$First = $True
	If($Script:MSHotfixes)
	{
		ForEach($Item in $Script:MSHotfixes)
		{
			If(!$First -and $Save -ne "$($Item.HotFixID)")
			{
				If($MSWord -or $PDF)
				{
					$WordTableRowHash = @{ 
					HotFixID    = "";
					ServerName  = "";
					Caption     = "";
					Description = "";
					InstalledBy = "";
					InstalledOn = "";}
					$ItemsWordTable += $WordTableRowHash;
				}
				If($Text)
				{
					Line 0 ""
				}
				If($HTML)
				{
					$rowdata += @(,(
					"",$htmlwhite,
					"",$htmlwhite,
					"",$htmlwhite,
					"",$htmlwhite,
					"",$htmlwhite,
					"",$htmlwhite))
				}
			}

			If($MSWord -or $PDF)
			{
				$WordTableRowHash = @{ 
				HotFixID    = $Item.HotFixID;
				ServerName  = $Item.ServerName;
				Caption     = $Item.Caption;
				Description = $Item.Description;
				InstalledBy = $Item.InstalledBy;
				InstalledOn = $Item.InstalledOn.ToString();}
				$ItemsWordTable += $WordTableRowHash;
			}
			If($Text)
			{
				Line 1 ( "{0,-25} {1,-15} {2,-45} {3,-20} {4,-35} {5,-22}" -f `
				$Item.HotFixID, $Item.ServerName, $Item.Caption, $Item.Description, $Item.InstalledBy, $Item.InstalledOn.ToString())
			}
			If($HTML)
			{
				$rowdata += @(,(
				$Item.HotFixID,$htmlwhite,
				$Item.ServerName,$htmlwhite,
				$Item.Caption,$htmlwhite,
				$Item.Description,$htmlwhite,
				$Item.InstalledBy,$htmlwhite,
				$Item.InstalledOn.ToString(),$htmlwhite))
			}

			$Save = "$($Item.HotFixID)"
			If($First)
			{
				$First = $False
			}
		}
	
		If($MSWord -or $PDF)
		{
			$Table = AddWordTable -Hashtable $ItemsWordTable `
			-Columns HotFixID, 
				ServerName, 
				Caption, 
				Description, 
				InstalledBy, 
				InstalledOn	`
			-Headers "Hotfix ID", 
				"Server Name", 
				"Caption", 
				"Description", 
				"Installed By", 
				"Installed On Date" `
			-AutoFit $wdAutoFitContent;

			SetWordCellFormat -Collection $Table -Size 9 -BackgroundColor $wdColorWhite
			SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			$ItemsWordTable = $Null
		}
		If($Text)
		{
			Line 0 ""
		}
		If($HTML)
		{
			$columnHeaders = @(
				"Hotfix ID", ($global:htmlsb),
				"Server Name", ($global:htmlsb),
				"Caption", ($global:htmlsb),
				"Description", ($global:htmlsb),
				"Installed By", ($global:htmlsb),
				"Installed On Date",($global:htmlsb))
							
			$msg = ""
			FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -tablewidth "600"
			WriteHTMLLine 0 0 " "
		}
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 ""
			WriteWordLine 0 0 "***No Microsoft Hotfixes and Updates were found***" "" $Null 0 $False $True
			WriteWordLine 0 0 ""
		}
		If($Text)
		{
			Line 0 ""
			Line 1 "***No Microsoft Hotfixes and Updates were found***"
			Line 0 ""
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 ""
			WriteHTMLLine 0 0 "***No Microsoft Hotfixes and Updates were found***" -Option $htmlBold
			WriteHTMLLine 0 0 ""
		}
	}

	Write-Verbose "$(Get-Date -Format G): Finished Creating Appendix M Microsoft Hotfixes and Updates"
	Write-Verbose "$(Get-Date -Format G): "
}
#endregion

#region appendix N
Function OutputAppendixN
{
	#added in V1.16
	Write-Verbose "$(Get-Date -Format G): Create Appendix N Windows Installed Components"

	$Script:WinInstalledComponents = $Script:WinInstalledComponents | Sort-Object DisplayName, Name, DDCName
	
	If($CSV)
	{
		$File = "$($Script:pwdpath)\$($Script:farm.FarmName)_HealthCheckV2_AppendixN_InstalledRolesandFeatures.csv"
		$Script:WinInstalledComponents | Export-CSV -Force -Encoding ASCII -NoTypeInformation -Path $File
	}
	
	If($MSWord -or $PDF)
	{
		$selection.InsertNewPage()
		WriteWordLine 1 0 "Appendix N - Windows Installed Components"
		[System.Collections.Hashtable[]] $ItemsWordTable = @();
	}
	If($Text)
	{
		Line 0 "Appendix N - Windows Installed Components"
	}
	If($HTML)
	{
		WriteHTMLLine 1 0 "Appendix N - Windows Installed Components"
		$rowdata = @()
	}

	If($Script:RunningOS -like "*2008*")
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "Unable to determine for a Server running Server 2008 or 2008 R2"
			WriteWordLine 0 0 ""
		}
		If($Text)
		{
			Line 1 "Unable to determine for a Server running Server 2008 or 2008 R2"
			Line 0 ""
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "Unable to determine for a Server running Server 2008 or 2008 R2"
			WriteHTMLLine 0 0 ""
		}
	}
	Else
	{
		If($Text)
		{
			Line 0 ""
			Line 1 "Display Name                                       Name                           Server Name      Feature Type   "
			Line 1 "=================================================================================================================="
			#       12345678901234567890123456789012345678901234567890S123456789012345678901234567890S123456789012345SS123456789012345
			#       Graphical Management Tools and Infrastructure      NET-Framework-45-Features      XXXXXXXXXXXXXXX  Role Service
			#       50                                                 30                             15               15
		}

		$Save = ""
		$First = $True
		If($Script:WinInstalledComponents)
		{
			ForEach($Item in $Script:WinInstalledComponents)
			{
				If(!$First -and $Save -ne "$($Item.DisplayName)$($Item.Name)")
				{
					If($MSWord -or $PDF)
					{
						$WordTableRowHash = @{ 
						DisplayName = "";
						Name        = "";
						ServerName  = "";
						FeatureType = "";}
						$ItemsWordTable += $WordTableRowHash;
					}
					If($Text)
					{
						Line 0 ""
					}
					If($HTML)
					{
						$rowdata += @(,(
						"",$htmlwhite,
						"",$htmlwhite,
						"",$htmlwhite,
						"",$htmlwhite))
					}
				}

				If($MSWord -or $PDF)
				{
					$WordTableRowHash = @{ 
					DisplayName = $Item.DisplayName;
					Name        = $Item.Name;
					ServerName  = $Item.ServerName;
					FeatureType = $Item.FeatureType;}
					$ItemsWordTable += $WordTableRowHash;
				}
				If($Text)
				{
					Line 1 ( "{0,-50} {1,-30} {2,-15}  {3,-15}" -f `
					$Item.DisplayName, $Item.Name, $Item.ServerName, $Item.FeatureType)
				}
				If($HTML)
				{
					$rowdata += @(,(
					$Item.DisplayName,$htmlwhite,
					$Item.Name,$htmlwhite,
					$Item.ServerName,$htmlwhite,
					$Item.FeatureType,$htmlwhite))
				}

				$Save = "$($Item.DisplayName)$($Item.Name)"
				If($First)
				{
					$First = $False
				}
			}

			If($MSWord -or $PDF)
			{
				$Table = AddWordTable -Hashtable $ItemsWordTable `
				-Columns DisplayName, 
					Name, 
					ServerName, 
					FeatureType	`
				-Headers "Display Name", 
					"Name", 
					"Server Name", 
					"Feature Type" `
				-AutoFit $wdAutoFitContent;

				SetWordCellFormat -Collection $Table -Size 9 -BackgroundColor $wdColorWhite
				SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				$ItemsWordTable = $Null
			}
			If($Text)
			{
				Line 0 ""
			}
			If($HTML)
			{
				$columnHeaders = @(
					"Display Name", ($global:htmlsb),
					"Name", ($global:htmlsb),
					"Server Name", ($global:htmlsb),
					"Feature Type",($global:htmlsb))
								
				$msg = ""
				FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -tablewidth "600"
				WriteHTMLLine 0 0 " "
			}
		}
		Else
		{
			If($MSWord -or $PDF)
			{
				WriteWordLine 0 0 ""
				WriteWordLine 0 0 "***No Windows installed Roles and Features were found***" "" $Null 0 $False $True
				WriteWordLine 0 0 ""
			}
			If($Text)
			{
				Line 0 ""
				Line 1 "***No Windows installed Roles and Features were found***"
				Line 0 ""
			}
			If($HTML)
			{
				WriteHTMLLine 0 0 ""
				WriteHTMLLine 0 0 "***No Windows installed Roles and Features were found***" -Option $htmlBold
				WriteHTMLLine 0 0 ""
			}
		}
	}

	Write-Verbose "$(Get-Date -Format G): Finished Creating Appendix N Windows Installed Components"
	Write-Verbose "$(Get-Date -Format G): "
}
#endregion

#region appendix O
Function OutputAppendixO
{
	#added in V1.16
	Write-Verbose "$(Get-Date -Format G): Create Appendix O PVS Processes"

	$Script:PVSProcessItems = $Script:PVSProcessItems | Sort-Object ProcessName, ServerName
	
	If($CSV)
	{
		$File = "$($Script:pwdpath)\$($Script:farm.FarmName)_HealthCheckV2_AppendixO_PVSProcesses.csv"
		$Script:PVSProcessItems | Export-CSV -Force -Encoding ASCII -NoTypeInformation -Path $File
	}
	
	If($MSWord -or $PDF)
	{
		$selection.InsertNewPage()
		WriteWordLine 1 0 "Appendix O - PVS Processes"
		[System.Collections.Hashtable[]] $ItemsWordTable = @();
	}
	If($Text)
	{
		Line 0 "Appendix O - PVS Processes"
		Line 0 ""
		Line 1 "Process Name  Server Name     Status     "
		Line 1 "========================================="
		#       1234567890123S123456789012345S12345678901
		#       StreamProcess XXXXXXXXXXXXXXX Not Running
		#       13            15              11
	}
	If($HTML)
	{
		WriteHTMLLine 1 0 "Appendix O - PVS Processes"
		$rowdata = @()
	}

	$Save = ""
	$First = $True
	If($Script:PVSProcessItems)
	{
		ForEach($Item in $Script:PVSProcessItems)
		{
			If(!$First -and $Save -ne "$($Item.ProcessName)")
			{
				If($MSWord -or $PDF)
				{
					$WordTableRowHash = @{ 
					ProcessName = "";
					ServerName  = "";
					Status      = "";}
					$ItemsWordTable += $WordTableRowHash;
				}
				If($Text)
				{
					Line 0 ""
				}
				If($HTML)
				{
					$rowdata += @(,(
					"",$htmlwhite,
					"",$htmlwhite,
					"",$htmlwhite))
				}
			}

			If($MSWord -or $PDF)
			{
				$WordTableRowHash = @{ 
				ProcessName = $Item.ProcessName;
				ServerName  = $Item.ServerName;
				Status      = $Item.Status;}
				$ItemsWordTable += $WordTableRowHash;
			}
			If($Text)
			{
				Line 1 ( "{0,-13} {1,-15} {2,-11}" -f `
				$Item.ProcessName, $Item.ServerName, $Item.Status)
			}
			If($HTML)
			{
				$rowdata += @(,(
				$Item.ProcessName,$htmlwhite,
				$Item.ServerName,$htmlwhite,
				$Item.Status,$htmlwhite))
			}

			$Save = "$($Item.ProcessName)"
			If($First)
			{
				$First = $False
			}
		}
	
		If($MSWord -or $PDF)
		{
			$Table = AddWordTable -Hashtable $ItemsWordTable `
			-Columns ProcessName, 
				ServerName, 
				Status `
			-Headers "Process Name", 
				"Server Name", 
				"Status" `
			-AutoFit $wdAutoFitContent;

			SetWordCellFormat -Collection $Table -Size 9 -BackgroundColor $wdColorWhite
			SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			$ItemsWordTable = $Null
		}
		If($Text)
		{
			Line 0 ""
		}
		If($HTML)
		{
			$columnHeaders = @(
				"Process Name", ($global:htmlsb),
				"Server Name", ($global:htmlsb),
				"Status",($global:htmlsb))
							
			$msg = ""
			FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -tablewidth "600"
			WriteHTMLLine 0 0 " "
		}
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 ""
			WriteWordLine 0 0 "***No PVS Processes were found***" "" $Null 0 $False $True
			WriteWordLine 0 0 ""
		}
		If($Text)
		{
			Line 0 ""
			Line 1 "***No PVS Processes were found***"
			Line 0 ""
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 ""
			WriteHTMLLine 0 0 "***No PVS Processes were found***" -Option $htmlBold
			WriteHTMLLine 0 0 ""
		}
	}

	Write-Verbose "$(Get-Date -Format G): Finished Creating Appendix O PVS Processes"
	Write-Verbose "$(Get-Date -Format G): "
}
#endregion

#region appendix p
Function OutputAppendixP
{
	#added in V1.23
	Write-Verbose "$(Get-Date -Format G): Create Appendix P Items to Review"

	$Script:ItemsToReview = $Script:ItemsToReview | Sort-Object ItemText
	
	If($CSV)
	{
		$File = "$($Script:pwdpath)\$($Script:farm.FarmName)_HealthCheckV2_AppendixP_ItemsToReview.csv"
		$Script:ItemsToReview | Export-CSV -Force -Encoding ASCII -NoTypeInformation -Path $File
	}
	
	If($MSWord -or $PDF)
	{
		$selection.InsertNewPage()
		WriteWordLine 1 0 "Appendix P - Items to Review"
		[System.Collections.Hashtable[]] $ItemsWordTable = @();
	}
	If($Text)
	{
		Line 0 "Appendix P - Items to Review"
		Line 0 ""
		Line 1 "Item                                   "
		Line 1 "======================================="
		#       123456789012345678901234567890134567890
		#       ItemText
	}
	If($HTML)
	{
		WriteHTMLLine 1 0 "Appendix P - Items to Review"
		$rowdata = @()
	}

	If($Script:ItemsToReview)
	{
		ForEach($Item in $Script:ItemsToReview)
		{
			If($MSWord -or $PDF)
			{
				$WordTableRowHash = @{ 
				ItemText = $Item.ItemText;}
				$ItemsWordTable += $WordTableRowHash;
			}
			If($Text)
			{
				Line 1 ( "{0,-40}" -f $Item.ItemText)
			}
			If($HTML)
			{
				$rowdata += @(,(
				$Item.ItemText,$htmlwhite))
			}
		}
	
		If($MSWord -or $PDF)
		{
			$Table = AddWordTable -Hashtable $ItemsWordTable `
			-Columns ItemText `
			-Headers "Item" `
			-AutoFit $wdAutoFitContent;

			SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			$ItemsWordTable = $Null
		}
		If($Text)
		{
			Line 0 ""
		}
		If($HTML)
		{
			$columnHeaders = @(
				"Item",($global:htmlsb))
							
			$msg = ""
			FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
			WriteHTMLLine 0 0 " "
		}
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 ""
			WriteWordLine 0 0 "***No Items to Review were found***" "" $Null 0 $False $True
			WriteWordLine 0 0 ""
		}
		If($Text)
		{
			Line 0 ""
			Line 1 "***No Items to Review were found***"
			Line 0 ""
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 ""
			WriteHTMLLine 0 0 "***No Items to Review were found***" -Option $htmlBold
			WriteHTMLLine 0 0 ""
		}
	}

	Write-Verbose "$(Get-Date -Format G): Finished Creating Appendix P Items to Review"
	Write-Verbose "$(Get-Date -Format G): "
}
#endregion

#region appendix Q
Function OutputAppendixQ
{
	#added in V1.23
	Write-Verbose "$(Get-Date -Format G): Create Appendix Q Server Items to Review"

	$Script:ServerComputerItemsToReview  = $Script:ServerComputerItemsToReview | Sort-Object ServerName
	$Script:ServerDriveItemsToReview     = $Script:ServerDriveItemsToReview | Sort-Object DriveCaption, ServerName
	$Script:ServerProcessorItemsToReview = $Script:ServerProcessorItemsToReview | Sort-Object ServerName
	$Script:ServerNICItemsToReview       = $Script:ServerNICItemsToReview | Sort-Object ServerName
	
	If($CSV)
	{
		$File = "$($Script:pwdpath)\$($Script:farm.FarmName)_HealthCheckV2_AppendixQ_ServerComputerItemsToReview.csv"
		$Script:ServerComputerItemsToReview | Export-CSV -Force -Encoding ASCII -NoTypeInformation -Path $File

		$File = "$($Script:pwdpath)\$($Script:farm.FarmName)_HealthCheckV2_AppendixQ_ServerDriveItemsToReview.csv"
		$Script:ServerDriveItemsToReview | Export-CSV -Force -Encoding ASCII -NoTypeInformation -Path $File

		$File = "$($Script:pwdpath)\$($Script:farm.FarmName)_HealthCheckV2_AppendixQ_ServerProcessorItemsToReview.csv"
		$Script:ServerProcessorItemsToReview | Export-CSV -Force -Encoding ASCII -NoTypeInformation -Path $File

		$File = "$($Script:pwdpath)\$($Script:farm.FarmName)_HealthCheckV2_AppendixQ_ServerNICItemsToReview.csv"
		$Script:ServerNICItemsToReview | Export-CSV -Force -Encoding ASCII -NoTypeInformation -Path $File
	}
	
	If($MSWord -or $PDF)
	{
		$selection.InsertNewPage()
		WriteWordLine 1 0 "Appendix Q - Server Items to Review"
		WriteWordLine 2 0 "Computer Items to Review"
		[System.Collections.Hashtable[]] $ItemsWordTable = @();
	}
	If($Text)
	{
		Line 0 "Appendix Q - Server Items to Review"
		Line 0 ""
		Line 1 "Computer Items to Review"
		Line 2 "Server Name     Operating System                        Power Plan        RAM   Physical  Logical"
		Line 2 "                                                                          (GB)  Procs     Procs  "
		Line 2 "================================================================================================="
		#       123456789012345S12345678901234567890123456789012345678SS1234567890123456SS1234SS12345678SS1234567
		#       XXXXXXXXXXXXXXX Microsoft Windows Server 2019 Standard High performance   9999  999       999
		#       15
	}
	If($HTML)
	{
		WriteHTMLLine 1 0 "Appendix Q - Server Items to Review"
		WriteHTMLLine 2 0 "Computer Items to Review"
		$rowdata = @()
	}

	If($Script:ServerComputerItemsToReview)
	{
		ForEach($Item in $Script:ServerComputerItemsToReview)
		{
			If($MSWord -or $PDF)
			{
				$WordTableRowHash = @{ 
				ServerName         = $Item.serverName;
				OperatingSystem    = $Item.OperatingSystem;
				PowerPlan          = $Item.PowerPlan;
				TotalRam           = $Item.TotalRam;
				PhysicalProcessors = $Item.PhysicalProcessors;
				LogicalProcessors  = $Item.LogicalProcessors;}
				$ItemsWordTable += $WordTableRowHash;
			}
			If($Text)
			{
				Line 2 ( "{0,-15} {1,-38}  {2,-16}  {3,4}  {4,8}  {5,7}" -f `
				$Item.ServerName, $Item.OperatingSystem, $Item.PowerPlan, $Item.TotalRam, $Item.PhysicalProcessors, $Item.LogicalProcessors)
			}
			If($HTML)
			{
				$rowdata += @(,(
				$Item.serverName,$htmlwhite,
				$Item.OperatingSystem,$htmlwhite,
				$Item.PowerPlan,$htmlwhite,
				$Item.TotalRam,$htmlwhite,
				$Item.PhysicalProcessors,$htmlwhite,
				$Item.LogicalProcessors,$htmlwhite))
			}
		}
	
		If($MSWord -or $PDF)
		{
			$Table = AddWordTable -Hashtable $ItemsWordTable `
			-Columns ServerName, 
				OperatingSystem, 
				PowerPlan, 
				TotalRam, 
				PhysicalProcessors, 
				LogicalProcessors	`
			-Headers "Server Name", 
				"Operating System", 
				"Power Plan", 
				"RAM (GB)", 
				"Physical Procs", 
				"Logical Procs" `
			-AutoFit $wdAutoFitContent;

			SetWordCellFormat -Collection $Table -Size 9 -BackgroundColor $wdColorWhite
			SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			$ItemsWordTable = $Null
		}
		If($Text)
		{
			Line 0 ""
		}
		If($HTML)
		{
			$columnHeaders = @(
				"Server Name", ($global:htmlsb),
				"Operating System", ($global:htmlsb),
				"Power Plan", ($global:htmlsb),
				"RAM (GB)", ($global:htmlsb),
				"Physical Procs", ($global:htmlsb),
				"Logical Procs",($global:htmlsb))
							
			$msg = ""
			FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -tablewidth "600"
			WriteHTMLLine 0 0 " "
		}
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 ""
			WriteWordLine 0 0 "***No Server Items to Review were found***" "" $Null 0 $False $True
			WriteWordLine 0 0 ""
		}
		If($Text)
		{
			Line 0 ""
			Line 2 "***No Server Items to Review were found***"
			Line 0 ""
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 ""
			WriteHTMLLine 0 0 "***No Server Items to Review were found***" -Option $htmlBold
			WriteHTMLLine 0 0 ""
		}
	}

	If($MSWord -or $PDF)
	{
		WriteWordLine 2 0 "Drive Items to Review"
		[System.Collections.Hashtable[]] $ItemsWordTable = @();
	}
	If($Text)
	{
		Line 1 "Drive Items to Review"
		Line 2 "Server Name     Caption  Size (GB)"
		Line 2 "=============================================="
		#       123456789012345S1234567SS123456789
		#       XXXXXXXXXXXXXXX C:            9999
		#       15
	}
	If($HTML)
	{
		WriteHTMLLine 2 0 "Drive Items to Review"
		$rowdata = @()
	}

	If($Script:ServerDriveItemsToReview)
	{
		ForEach($Item in $Script:ServerDriveItemsToReview)
		{
			If($MSWord -or $PDF)
			{
				$WordTableRowHash = @{ 
				ServerName   = $Item.serverName;
				DriveCaption = $Item.DriveCaption;
				DriveSize    = $Item.DriveSize;}
				$ItemsWordTable += $WordTableRowHash;
			}
			If($Text)
			{
				Line 2 ( "{0,-15} {1,-7}  {2,9}" -f `
				$Item.ServerName, $Item.DriveCaption, $Item.DriveSize)
			}
			If($HTML)
			{
				$rowdata += @(,(
				$Item.serverName,$htmlwhite,
				$Item.DriveCaption,$htmlwhite,
				$Item.DriveSize,$htmlwhite))
			}
		}
	
		If($MSWord -or $PDF)
		{
			$Table = AddWordTable -Hashtable $ItemsWordTable `
			-Columns ServerName, 
				ThreadsPerPort, 
				EnableNonBlockingIO	`
			-Headers "Server Name", 
				"Caption", 
				"Size (GB)" `
			-AutoFit $wdAutoFitContent;

			SetWordCellFormat -Collection $Table -Size 9 -BackgroundColor $wdColorWhite
			SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			$ItemsWordTable = $Null
		}
		If($Text)
		{
			Line 0 ""
		}
		If($HTML)
		{
			$columnHeaders = @(
				"Server Name", ($global:htmlsb),
				"Caption", ($global:htmlsb),
				"Size (GB)",($global:htmlsb))
							
			$msg = ""
			FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -tablewidth "600"
			WriteHTMLLine 0 0 " "
		}
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 ""
			WriteWordLine 0 0 "***No Drive Items to Review were found***" "" $Null 0 $False $True
			WriteWordLine 0 0 ""
		}
		If($Text)
		{
			Line 0 ""
			Line 2 "***No Drive Items to Review were found***"
			Line 0 ""
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 ""
			WriteHTMLLine 0 0 "***No Drive Items to Review were found***" -Option $htmlBold
			WriteHTMLLine 0 0 ""
		}
	}

	If($MSWord -or $PDF)
	{
		WriteWordLine 2 0 "Processor Items to Review"
		[System.Collections.Hashtable[]] $ItemsWordTable = @();
	}
	If($Text)
	{
		Line 1 "Processor Items to Review"
		Line 2 "Server Name     Cores  Logical Procs"
		Line 2 "===================================="
		#       123456789012345S12345SS1234567890123
		#       XXXXXXXXXXXXXXX  9999           9999
		#       15
	}
	If($HTML)
	{
		WriteHTMLLine 2 0 "Processor Items to Review"
		$rowdata = @()
	}

	If($Script:ServerProcessorItemsToReview)
	{
		ForEach($Item in $Script:ServerProcessorItemsToReview)
		{
			If($MSWord -or $PDF)
			{
				$WordTableRowHash = @{ 
				ServerName   = $Item.serverName;
				Cores        = $Item.Cores;
				LogicalProcs = $Item.LogicalProcs;}
				$ItemsWordTable += $WordTableRowHash;
			}
			If($Text)
			{
				Line 2 ( "{0,-15} {1,5}  {2,13}" -f `
				$Item.ServerName, $Item.Cores , $Item.LogicalProcs)
			}
			If($HTML)
			{
				$rowdata += @(,(
				$Item.serverName,$htmlwhite,
				$Item.Cores,$htmlwhite,
				$Item.LogicalProcs,$htmlwhite))
			}
		}
	
		If($MSWord -or $PDF)
		{
			$Table = AddWordTable -Hashtable $ItemsWordTable `
			-Columns ServerName, 
				Cores, 
				LogicalProcs `
			-Headers "Server Name", 
				"Cores", 
				"Logical Procs" `
			-AutoFit $wdAutoFitContent;

			SetWordCellFormat -Collection $Table -Size 9 -BackgroundColor $wdColorWhite
			SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			$ItemsWordTable = $Null
		}
		If($Text)
		{
			Line 0 ""
		}
		If($HTML)
		{
			$columnHeaders = @(
				"Server Name", ($global:htmlsb),
				"Cores", ($global:htmlsb),
				"Logical Procs",($global:htmlsb))
							
			$msg = ""
			FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -tablewidth "600"
			WriteHTMLLine 0 0 " "
		}
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 ""
			WriteWordLine 0 0 "***No Processor Items to Review were found***" "" $Null 0 $False $True
			WriteWordLine 0 0 ""
		}
		If($Text)
		{
			Line 0 ""
			Line 2 "***No Processor Items to Review were found***"
			Line 0 ""
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 ""
			WriteHTMLLine 0 0 "***No Processor Items to Review were found***" -Option $htmlBold
			WriteHTMLLine 0 0 ""
		}
	}

	If($MSWord -or $PDF)
	{
		WriteWordLine 2 0 "NIC Items to Review"
		[System.Collections.Hashtable[]] $ItemsWordTable = @();
	}
	If($Text)
	{
		Line 1 "NIC Items to Review"
		Line 2 "Server Name     NIC Name                                    Manufacturer          Power Mgmt  RSS     "
		Line 2 "======================================================================================================"
		#       123456789012345S123456789012345678901234567890123456789012SS12345678901234567890SS1234567890SS12345678
		#       XXXXXXXXXXXXXXX Intel(R) 82574L Gigabit Network Connection  Intel Corporation     Disabled    Disabled
		#       15              42                                          20                    9           8
	}
	If($HTML)
	{
		WriteHTMLLine 2 0 "NIC Items to Review"
		$rowdata = @()
	}

	If($Script:ServerNICItemsToReview)
	{
		ForEach($Item in $Script:ServerNICItemsToReview)
		{
			If($MSWord -or $PDF)
			{
				$WordTableRowHash = @{ 
				ServerName   = $Item.serverName;
				Name         = $Item.Name;
				Manufacturer = $Item.Manufacturer;
				PowerMgmt    = $Item.PowerMgmt;
				RSS          = $Item.RSS;}
				$ItemsWordTable += $WordTableRowHash;
			}
			If($Text)
			{
				Line 2 ( "{0,-15} {1,-42}  {2,-20}  {3,-10}  {4,-8}" -f `
				$Item.ServerName, $Item.Name, $Item.Manufacturer, $Item.PowerMgmt, $Item.RSS)
			}
			If($HTML)
			{
				$rowdata += @(,(
				$Item.serverName,$htmlwhite,
				$Item.Name,$htmlwhite,
				$Item.Manufacturer,$htmlwhite,
				$Item.PowerMgmt,$htmlwhite,
				$Item.RSS,$htmlwhite))
			}
		}
	
		If($MSWord -or $PDF)
		{
			$Table = AddWordTable -Hashtable $ItemsWordTable `
			-Columns ServerName, 
				Name, 
				Manufacturer, 
				PowerMgmt, 
				RSS	`
			-Headers "Server Name", 
				"NIC Name", 
				"Manufacturer", 
				"Power Mgmt", 
				"RSS" `
			-AutoFit $wdAutoFitContent;

			SetWordCellFormat -Collection $Table -Size 9 -BackgroundColor $wdColorWhite
			SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			$ItemsWordTable = $Null
		}
		If($Text)
		{
			Line 0 ""
		}
		If($HTML)
		{
			$columnHeaders = @(
				"Server Name", ($global:htmlsb),
				"NIC Name", ($global:htmlsb),
				"Manufacturer", ($global:htmlsb),
				"Power Mgmt", ($global:htmlsb),
				"RSS",($global:htmlsb))
							
			$msg = ""
			FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -tablewidth "600"
			WriteHTMLLine 0 0 " "
		}
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 ""
			WriteWordLine 0 0 "***No NIC Items to Review were found***" "" $Null 0 $False $True
			WriteWordLine 0 0 ""
		}
		If($Text)
		{
			Line 0 ""
			Line 1 "***No NIC Items to Review were found***"
			Line 0 ""
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 ""
			WriteHTMLLine 0 0 "***No NIC Items to Review were found***" -Option $htmlBold
			WriteHTMLLine 0 0 ""
		}
	}

	Write-Verbose "$(Get-Date -Format G): Finished Creating Appendix Q Server Items to Review"
	Write-Verbose "$(Get-Date -Format G): "
}
#endregion

#region Appendixr
Function OutputAppendixR
{
	#added in V1.24
	Write-Verbose "$(Get-Date -Format G): Create Appendix R Citrix Installed Components"

	$Script:CtxInstalledComponents = $Script:CtxInstalledComponents | Sort-Object DisplayName, PVSServerName
	
	If($CSV)
	{
		$File = "$($Script:pwdpath)\$($Script:farm.FarmName)_HealthCheckV2_AppendixR_CitrixInstalledComponents.csv"
		$Script:CtxInstalledComponents | Export-CSV -Force -Encoding ASCII -NoTypeInformation -Path $File *> $Null
	}
	
	If($MSWord -or $PDF)
	{
		$selection.InsertNewPage()
		WriteWordLine 1 0 "Appendix R - Citrix Installed Components"
		[System.Collections.Hashtable[]] $ItemsWordTable = @();
	}
	If($Text)
	{
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
	}
	If($HTML)
	{
		WriteHTMLLine 1 0 "Appendix R - Citrix Installed Components"
		$rowdata = @()
	}
	
	$Save = ""
	$First = $True
	If($Script:CtxInstalledComponents)
	{
		ForEach($Item in $Script:CtxInstalledComponents)
		{
			If(!$First -and $Save -ne "$($Item.DisplayName)$($Item.DisplayVersion)")
			{
				If($MSWord -or $PDF)
				{
					$WordTableRowHash = @{ 
					DisplayName    = "";
					DisplayVersion = "";
					PVSServerName  = "";}
					$ItemsWordTable += $WordTableRowHash;
				}
				If($Text)
				{
					Line 0 ""
				}
				If($HTML)
				{
					$rowdata += @(,(
					"",$htmlwhite,
					"",$htmlwhite,
					"",$htmlwhite))
				}
			}

			If($MSWord -or $PDF)
			{
				$WordTableRowHash = @{ 
				DisplayName    = $Item.DisplayName;
				DisplayVersion = $Item.DisplayVersion;
				PVSServerName  = $Item.PVSServerName;}
				$ItemsWordTable += $WordTableRowHash;
			}
			If($Text)
			{
				Line 1 ( "{0,$NegativeMaxLength} {1,-25} {2,-40}" -f `
				$Item.DisplayName, $Item.DisplayVersion, $Item.PVSServerName)
			}
			If($HTML)
			{
				$rowdata += @(,(
				$Item.DisplayName,$htmlwhite,
				$Item.DisplayVersion,$htmlwhite,
				$Item.PVSServerName,$htmlwhite))
			}
			
			$Save = "$($Item.DisplayName)$($Item.DisplayVersion)"
			If($First)
			{
				$First = $False
			}
		}
	
		If($MSWord -or $PDF)
		{
			$Table = AddWordTable -Hashtable $ItemsWordTable `
			-Columns DisplayName, 
				DisplayVersion, 
				PVSServerName	`
			-Headers "Display Name", 
				"Display Version", 
				"PVS Server Name" `
			-AutoFit $wdAutoFitContent;

			SetWordCellFormat -Collection $Table -Size 9 -BackgroundColor $wdColorWhite
			SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			$ItemsWordTable = $Null
		}
		If($Text)
		{
			Line 0 ""
		}
		If($HTML)
		{
			$columnHeaders = @(
				"Display Name", ($global:htmlsb),
				"Display Version", ($global:htmlsb),
				"PVS Server Name",($global:htmlsb))
							
			$msg = ""
			FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -tablewidth "600"
			WriteHTMLLine 0 0 " "
		}
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 ""
			WriteWordLine 0 0 "***No Citrix Installed Components were found***" "" $Null 0 $False $True
			WriteWordLine 0 0 ""
		}
		If($Text)
		{
			Line 0 ""
			Line 1 "***No Citrix Installed Components were found***"
			Line 0 ""
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 ""
			WriteHTMLLine 0 0 "***No Citrix Installed Components were found***" -Option $htmlBold
			WriteHTMLLine 0 0 ""
		}
	}

	Write-Verbose "$(Get-Date -Format G): Finished Creating Appendix R Citrix Installed Components"
	Write-Verbose "$(Get-Date -Format G): "
}
#endregion

#region script end
Function ProcessScriptEnd
{
	Write-Verbose "$(Get-Date -Format G): Script has completed"
	Write-Verbose "$(Get-Date -Format G): "

	#http://poshtips.com/measuring-elapsed-time-in-powershell/
	Write-Verbose "$(Get-Date -Format G): Script started: $($Script:StartTime)"
	Write-Verbose "$(Get-Date -Format G): Script ended: $(Get-Date)"
	$runtime = $(Get-Date) - $Script:StartTime
	$Str = [string]::format("{0} days, {1} hours, {2} minutes, {3}.{4} seconds",
		$runtime.Days,
		$runtime.Hours,
		$runtime.Minutes,
		$runtime.Seconds,
		$runtime.Milliseconds)
	Write-Verbose "$(Get-Date -Format G): Elapsed time: $($Str)"

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
		$SIFile = "$($Script:pwdpath)\PVSInventoryScriptInfo_$(Get-Date -f yyyy-MM-dd_HHmm).txt"
		Out-File -FilePath $SIFile -InputObject "" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Add DateTime        : $AddDateTime" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "AdminAddress        : $AdminAddress" 4>$Null
		If($MSWORD -or $PDF)
		{
			Out-File -FilePath $SIFile -Append -InputObject "Company Name        : $Script:CoName" 4>$Null		
			Out-File -FilePath $SIFile -Append -InputObject "Company Address     : $CompanyAddress" 4>$Null		
			Out-File -FilePath $SIFile -Append -InputObject "Company Email       : $CompanyEmail" 4>$Null		
			Out-File -FilePath $SIFile -Append -InputObject "Company Fax         : $CompanyFax" 4>$Null		
			Out-File -FilePath $SIFile -Append -InputObject "Company Phone       : $CompanyPhone" 4>$Null		
			Out-File -FilePath $SIFile -Append -InputObject "Cover Page          : $CoverPage" 4>$Null
		}
		Out-File -FilePath $SIFile -Append -InputObject "Dev                 : $Dev" 4>$Null
		If($Dev)
		{
			Out-File -FilePath $SIFile -Append -InputObject "DevErrorFile        : $Script:DevErrorFile" 4>$Null
		}
		Out-File -FilePath $SIFile -Append -InputObject "Domain              : $Domain" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "EndDate             : $EndDate" 4>$Null
		If($HTML)
		{
			Out-File -FilePath $SIFile -Append -InputObject "HTMLFilename        : $Script:HTMLFileName" 4>$Null
		}
		If($MSWord)
		{
			Out-File -FilePath $SIFile -Append -InputObject "WordFilename        : $Script:WordFileName" 4>$Null
		}
		If($PDF)
		{
			Out-File -FilePath $SIFile -Append -InputObject "PDFFilename         : $Script:PDFFileName" 4>$Null
		}
		If($Text)
		{
			Out-File -FilePath $SIFile -Append -InputObject "TextFilename        : $Script:TextFileName" 4>$Null
		}
		Out-File -FilePath $SIFile -Append -InputObject "Folder              : $Folder" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "From                : $From" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Log                 : $Log" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Report Footer       : $ReportFooter" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Save As HTML        : $HTML" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Save As PDF         : $PDF" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Save As TEXT        : $TEXT" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Save As WORD        : $MSWORD" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Script Info         : $ScriptInfo" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Smtp Port           : $SmtpPort" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Smtp Server         : $SmtpServer" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Start Date          : $StartDate" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Title               : $Script:Title" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "To                  : $To" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Use SSL             : $UseSSL" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "User                : $User" 4>$Null
		If($MSWORD -or $PDF)
		{
			Out-File -FilePath $SIFile -Append -InputObject "Username            : $UserName" 4>$Null
		}
		Out-File -FilePath $SIFile -Append -InputObject "" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "OS Detected         : $Script:RunningOS" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "PoSH version        : $($Host.Version)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "PSCulture           : $PSCulture" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "PSUICulture         : $PSUICulture" 4>$Null
		If($MSWORD -or $PDF)
		{
			Out-File -FilePath $SIFile -Append -InputObject "Word language       : $Script:WordLanguageValue" 4>$Null
			Out-File -FilePath $SIFile -Append -InputObject "Word version        : $Script:WordProduct" 4>$Null
		}
		Out-File -FilePath $SIFile -Append -InputObject "" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Script start        : $Script:StartTime" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Elapsed time        : $Str" 4>$Null
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

	$ErrorActionPreference = $SaveEAPreference
}
#endregion

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

#script begins

#region script start Function
Function ProcessScriptStart
{
	$script:startTime = Get-Date
}
#endregion

#region script setup Function
Function ProcessScriptSetup
{
	CheckOnPoSHPrereqs

	SetupRemoting

	VerifyPVSServices

	GetPVSVersion

	GetPVSFarm
}
#endregion

ProcessScriptStart

ProcessScriptSetup

SetFileNames "$($Script:farm.FarmName)_HealthCheckV2"

ProcessPVSFarm

ProcessPVSSite

ProcessvDisksinFarm

ProcessStores

OutputAppendixA	#Appendix A - Advanced Server Items (Server/Network)

OutputAppendixB	#Appendix B - Advanced Server Items (Pacing/Device)

OutputAppendixC	#Appendix C - Configuration Wizard Settings

OutputAppendixD	#Appendix D - Server Bootstrap Items

OutputAppendixE	#Appendix E - DisableTaskOffload Settings

OutputAppendixF1	#Appendix F - Server PVS Service Items

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

#region finish script
Write-Verbose "$(Get-Date -Format G): Finishing up document"
#end of document processing

If(($MSWORD -or $PDF) -and ($Script:CoverPagesExist))
{
	$AbstractTitle = "PVS Health Check Report for Farm $($Script:farm.FarmName)"
	$SubjectTitle = "PVS Health Check Report for Farm $($Script:farm.FarmName)"
	UpdateDocumentProperties $AbstractTitle $SubjectTitle
}

If($ReportFooter)
{
	OutputReportFooter
}

ProcessDocumentOutput

ProcessScriptEnd
#endregion
