# Citrix-PVS-Health-Check-2
# Requires PowerShell Version 3.0 or later
Creates a basic health check of a Citrix PVS 6.x or later farm in plain text, HTML, Word, or PDF.

This script was tested with PVS 6.1 running on Windows Server 2008 R2 and PVS 2112 running on Windows Server 2022.

Creates a document named after the PVS farm.

Version 2.0 changes the default output report from text to HTML.

The script must run from an elevated PowerShell session.

NOTE: The account used to run this script must have at least Read access to the SQL Server that holds the Citrix Provisioning databases.

This script is written using the old string-based crappy PowerShell because it supports PVS 6.x.
