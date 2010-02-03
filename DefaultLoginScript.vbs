'region Script Settings
'<ScriptSettings xmlns="http://tempuri.org/ScriptSettings.xsd">
'  <ScriptPackager>
'    <process>cscript.exe</process>
'    <arguments />
'    <extractdir>%TEMP%</extractdir>
'    <files />
'    <usedefaulticon>true</usedefaulticon>
'    <showinsystray>false</showinsystray>
'    <altcreds>false</altcreds>
'    <efs>true</efs>
'    <ntfs>true</ntfs>
'    <local>false</local>
'    <abortonfail>true</abortonfail>
'    <product />
'    <version>1.0.0.1</version>
'    <versionstring />
'    <comments />
'    <includeinterpreter>false</includeinterpreter>
'    <forcecomregistration>false</forcecomregistration>
'    <consolemode>false</consolemode>
'    <EnableChangelog>false</EnableChangelog>
'    <AutoBackup>false</AutoBackup>
'    <snapinforce>false</snapinforce>
'    <snapinshowprogress>false</snapinshowprogress>
'    <snapinautoadd>0</snapinautoadd>
'    <snapinpermanentpath />
'  </ScriptPackager>
'</ScriptSettings>
'endregion

'#Windows Script Host Script#

'### Name ###
'
'		DefaultLoginScript.vbs

'### Description ###
'
'		This script is used to perform various procedures at logon.  It makes use of a central
'		script library (ScriptLibrary.vbs) to abstract the more detailed implementation of common
'		procedures.



'region Script Source Code
Option Explicit
Const SCRIPTSERVER = "ADS1"
Const FILESERVER = "OSX-SERVER1"
Const FQDOMAINNAME = "HOME.LAN"

'region Implements Central Script Library (ScriptLibrary.vbs)
Function ScriptLibrary(procedureNameWithArgsSpaceDelimited)
Dim Shell
Set Shell = CreateObject("WScript.Shell")
Shell.Run("cscript \\" & SCRIPTSERVER & "\SYSVOL\" & FQDOMAINNAME & "\scripts\ScriptLibrary.vbs" & " " & procedureNameWithArgsSpaceDelimited)
End Function

'endregion

'Add Network Printer
ScriptLibrary "AddNetworkPrinter \\" & SCRIPTSERVER & "\CanoniP4500"
ScriptLibrary "AddNetworkPrinter \\" & SCRIPTSERVER & "\CanonMF4690"

'Set Default Printer
ScriptLibrary "SetNetworkPrinterAsDefault ""\\" & SCRIPTSERVER & "\CanonMF4690"""

'Map Apps & Shared Network Drives
'ScriptLibrary "MapNetworkDrive I: \\" & FILESERVER & "\NetApps False"
ScriptLibrary "MapNetworkDrive S: \\" & FILESERVER & "\Public False"

'Rename the Drives
'ScriptLibrary "RenameDrive I: ""Applications"""
ScriptLibrary "RenameDrive S: ""Shared Disk"""

'Make A Personal Downloads Folder if not already present
ScriptLibrary "CreateFolder \\" & FILESERVER & "\Users\%username%\Downloads"
'Create a shortcut for it in the main Personal Folder (My Documents)
ScriptLibrary "CreateShortcut Downloads ""\\" & FILESERVER & "\Users\%username%\Documents\Downloads.lnk"" \\" & FILESERVER & "\Users\%username%\Downloads"

'Make A "My Videos" Folder if not already present
ScriptLibrary "CreateFolder ""\\" & FILESERVER & "\Users\%username%\Documents\My Videos"""

'Redirect Favorites Folder
'ScriptLibrary "RedirectShellFolder Favorites \\" & FILESERVER & "\Users\%username%\Favorites"
'endregion