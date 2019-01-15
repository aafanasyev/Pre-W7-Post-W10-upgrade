# Pre-W7-Post-W10-upgrade
This repository contains two PowerShell scripts: pre (W7) and post(W10) upgrade.


#Pre-upgrade script 

It is used on Windows 7  to check some compliences that are useful in Active Directory environment. It contains checks of:
- Check to which collection machine is assigned to.
- Version of SCCM client.

#Post-upgrade script
It used on Windows 10 to provide some acctions. 
- Remove Windows 10 Built-In application as they defined in Excel or CSV file.
- Excel might be not installed on a machine. Excel file need to be converted in to CSV manually.
- Check if machine assigned to a specific AD collection.

**NOTE** It is possible to convert an Excel file in to CSV using PowerShell. It can be done even on machine where MS Office 2010+ is not installed. Read data from an Excel files requires a OLE DB. It is installed by default on Windows 10. However, Excel data can be stored in two types of file extentions "xls" or "xlsx". 

For the first a "Microsoft.Jet.OLEDB.4.0" driver is needed. It is an old driver and works only in 32 bits PowerShell. For the second a "Microsoft.ACE.OLEDB.12.0" driver is needed. This driver is not installed by default and requires 2007 Office System Driver: Data Connectivity Components and Microsoft Access Database Engine 2010 Redistributable. This overcomplicates Post upgrade process so it is not used in a script.
