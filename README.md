# Citrix-PVS-Health-Check-2
Creates a basic health check of a Citrix PVS 5.x or later farm in plain text, HTML, Word, or PDF.

This script has been tested with PVS 2112.
	
Creates a document named after the PVS farm.

The script must run from an elevated PowerShell session.
	
NOTE: The account used to run this script must have at least Read access to the SQL Server that holds the Citrix Provisioning databases.
	
This script is written using the old string-based crappy PowerShell because it supports PVS 5.x.
	
In order for the email stuff to work, this script requires at least PowerShell V3.
