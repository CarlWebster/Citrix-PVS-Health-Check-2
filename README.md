# Citrix-PVS-Health-Check-2
Creates a basic health check of a Citrix PVS 6.x or later farm in plain text, 
HTML, Word, or PDF.

This script has been tested with PVS 6.1 and 2112.

Creates a document named after the PVS farm.

Version 2.0 changes the default output report from text to HTML.

The script must run from an elevated PowerShell session.

NOTE: The account used to run this script must have at least Read access to the SQL 
Server that holds the Citrix Provisioning databases.

This script is written using the old string-based crappy PowerShell because it 
supports PVS 6.x.

For the email stuff to work, this script requires at least PowerShell V3.
