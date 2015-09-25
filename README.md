# Very Basic Script Editor

This is a simple editor for VBScript and JScript, with the ability to run the scripts from the IDE.

## Installation

 1. Extract the zip file to a directory of your choice
 2. Right click on VBSE.exe, select "Properties", click the "Unblock" button if present, and click "Ok"
 3. Right click on script56.chm, select "Properties", click the "Unblock" button if present, and click "Ok"
 4. If you are running Windows Vista or later, right click on VBSE.exe, and select "Run As Administrator". You may want to set up a shortcut to always run as administrator for VBSE

## Upgrading

If you are upgrading, you generally only need to replace VBSE.exe. If you are upgrading from 1.0, 1.1, 1.2, or 1.3, you should extract script56.chm into the program directory, and unblock it (see Installation step 3 for details).

## Scripting Environment

I have included several VB objects in the scripting environment, and am in the process of implementing DOM objects in the editor for scripting. These are accessible from both VBScript and JScript, but do remember that JScript is case-sensitive, while VBScript is not.

### DOM Objects

 * console - only log is officially supported, with unsupported access to the Console window
 * navigator - a simple implementation
 * window - only alert, confirm, and prompt are currently implemented

### Visual Basic Objects

 * App (referring to the IDE Program)
 * Clipboard
 * Me (referring to the IDE Window)
 * Printer
 * Printers
 * Screen

## Requirements

 * Administrator permissions
 * Windows NT 4.0 or higher
 * Microsoft Visual Basic 6.0 SP6 Runtime (included with Windows XP or higher) http://www.microsoft.com/en-us/download/details.aspx?id=24417
 * Microsoft Script Control (included with Windows 2000 or higher) http://www.microsoft.com/downloads/details.aspx?FamilyId=D7E31492-2595-49E6-8C02-1426FEC693AC
 * Windows Script 5.6 Documentation (included) https://www.microsoft.com/en-nz/download/details.aspx?id=2764

## Download

You can download the latest version from https://github.com/pmachapman/VBSE/releases

## Licence

This program is licensed under the GPLv3, except for the CommonDialog directory.

This program comes without any warranty. Use at your own risk!
