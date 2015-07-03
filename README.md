# Very Basic Script Editor

This is a simple editor for VBScipt and JScript, with the ability to run the scripts from the IDE.

## Scripting Environment

I have included several VB objects in the scripting environment, and am in the process of implementing DOM objects in the editor for scripting. These are accessible form both VBScript and JScript, but do remember that JScript is case-sensitive, while VBScript is not.

### DOM Objects

 * console - only log is supported
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
 * Microsoft Visual Basic 4.0 Runtime (included) https://support.microsoft.com/en-us/kb/196286
 * Microsoft Script Control http://www.microsoft.com/downloads/details.aspx?FamilyId=D7E31492-2595-49E6-8C02-1426FEC693AC&displaylang=en
 * Microsoft Common Dialog Control 6.0 (included)

## Download

You can download the executable from https://archive.azurewebsites.net/Programs/VBSE-1.2.zip

## This program comes without any warranty

As this is a program for programmers, you should only use if you know what you are doing! Be careful where you first run the program, and make sure you always run it as Administrator. It will automatically register COMDLG32.OCX in the program directory. If you do not want this to happen, you should relocate the VB40032.DLL and COMDLG32.OCX files to your SYSTEM32 or SYSWOW64 directories, and run REGSVR32 on COMDLG32.OCX.

If there is no MSSCRIPT.OCX file in your SYSTEM32 or SYSWOW64 directories (there usually will be), please install the Microsoft Script Control from http://www.microsoft.com/downloads/details.aspx?FamilyId=D7E31492-2595-49E6-8C02-1426FEC693AC&displaylang=en

Finally, use at your own risk!
