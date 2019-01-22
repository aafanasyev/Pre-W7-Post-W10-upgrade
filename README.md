# Pre-W7-Post-W10-upgrade
This repository contains two PowerShell scripts: pre (W7) and post(W10) upgrade.


# Pre-upgrade script 

It is used on Windows 7  to check some compliances that are useful in the Active Directory environment. It contains checks of:
- Check to which collection machine is assigned to.
- A version of SCCM client.

## Important modules:
To get computer information:

`Get-WmiObject -Class Win32_ComputerSystem` 

`Get-WmiObject -Class Win32_OperatingSystem`

To get AD information:

`New-Object System.DirectoryServices.DirectorySearcher`

`New-Object System.DirectoryServices.DirectoryEntry`

To get SMMC collection:

`Get-CMCollection`

`Get-CMCollectionMember`

SCCM Client Version

`Get-WMIObject -Namespace root\ccm -Class SMS_Client`

User Profiles to backup

Get-WMIObject -Class Win32_UserProfile

Windows Update

`New-Object -COM Microsoft.Update.Session`

Storage Volumes

`Get-Volume`

Software Compatibility

`---`

Snapshots

`Get-ItemProperty HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*`

`Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*`

NetworkShares

`Get-WmiObject -Class Win32_Share`

`Get-WmiObject -Class Win32_LogicalShareSecuritySetting`

Printers

`Get-WMIObject -Class Win32_Printer`

Convert output to HTML:

`ConvertTo-HTML`


# Post-upgrade script
It used on Windows 10 to provide following actions:
- Remove Windows 10 Built-In application as they defined in Excel or CSV file.
- Excel might be not installed on a machine. Excel file needs to be converted into CSV manually.
- Check if machine assigned to a specific AD collection.

**NOTE** It is possible to convert an Excel file into CSV using PowerShell. It can be done even on a machine where MS Office 2007+ is not installed. Read data from Excel files requires an OLE DB. It is installed by default on Windows 10. However, Excel data can be stored in two types of file extensions ".xls" or ".xlsx". 

For the first a "Microsoft.Jet.OLEDB.4.0" driver is needed. It is an old driver and works only in 32 bits PowerShell. For the second a "Microsoft.ACE.OLEDB.12.0" driver is needed. This driver is not installed by default and requires 2007 Office System Driver: Data Connectivity Components and Microsoft Access Database Engine 2010 Redistributable. This overcomplicates Post upgrade process so it is not used in a script.


## Important modules:
