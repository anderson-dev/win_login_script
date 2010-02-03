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
'		ScriptLibrary.vbs

'### Synopsis ###
'
'		ScriptLibrary.vbs nameOfProcedure [requiredArgumentList]

'### Description ###
'
'		This script is used as a central script library to abstract actual implementation from
'		the main (calling) script.  From within any script, a call to this script using the
'		WshShell object and proper syntax will perform the named Procedure using the appropriate
'		arguments if required.
'
'		The components are as follows:
'
'		nameOfProcedure			This is the full name of a Function contained within this script.
'		(Required Component)		
'
'		requiredArgumentList		A named Procedure may require zero or more arguments.  This is a 
'										list of arguments required by the named Procedure In the
'										form of: argument1 argument2 argument3 ...
'										where each argument is in the required order of the function's
'										signiature and only a space seperating each consecutive argument.


'### List Of Procedure Names ###
'

'region Script Source Code
Option Explicit
Const SCRIPTSERVER = "ADS1"
Const FQDOMAINNAME = "HOME.LAN"

'region EnvVarsInterpreter
Function EnvVarsInterpreter(stringToEvaluate)
Dim shell, regEx
Set shell = CreateObject("WScript.Shell")
Set regEx = New RegExp

'Substitute "%username%" for actual username
With regEx
.Pattern = "%username%"
.IgnoreCase = True
.Global = True
End With
stringToEvaluate = regEx.Replace(stringToEvaluate, Shell.ExpandEnvironmentStrings("%USERNAME%"))

'Substitute "%userprofile%" for actual user's profile path
With regEx
.Pattern = "%userprofile%"
.IgnoreCase = True
.Global = True
End With
stringToEvaluate = regEx.Replace(stringToEvaluate, Shell.ExpandEnvironmentStrings("%USERPROFILE%"))

'return modified string
EnvVarsInterpreter = stringToEvaluate
End Function
'endregion

'region Function to execute Procedures

'	Determine And Execute A Named Procedure
Function ExecuteProcedure(procedureName, argc)
'Determine which Procedure was entered by user
Dim result
Select Case procedureName
Case "CreateFile"
	CreateFile WScript.Arguments.Item(1), WScript.Arguments.Item(2), WScript.Arguments.Item(3)
Case "AppendFile"
	AppendFile WScript.Arguments.Item(1), WScript.Arguments.Item(2)
Case "DeleteFiles"
	DeleteFiles WScript.Arguments.Item(1)
Case "CreateFolder"
	CreateFolder WScript.Arguments.Item(1)
Case "DeleteFolder"
	DeleteFolder WScript.Arguments.Item(1)
Case "CopyFiles"
	CopyFiles WScript.Arguments.Item(1), WScript.Arguments.Item(2), WScript.Arguments.Item(3)
Case "MoveFiles"
	MoveFiles WScript.Arguments.Item(1), WScript.Arguments.Item(2)
Case "CopyFolder"
	CopyFolder WScript.Arguments.Item(1), WScript.Arguments.Item(2), WScript.Arguments.Item(3)
Case "MoveFolder"
	MoveFolder WScript.Arguments.Item(1), WScript.Arguments.Item(2)
Case "RenameDrive"
	RenameDrive WScript.Arguments.Item(1), WScript.Arguments.Item(2)
Case "CreateShortcut"
	CreateShortcut WScript.Arguments.Item(1), WScript.Arguments.Item(2), WScript.Arguments.Item(3)
Case "Run"
	If WScript.Arguments.Count > 2 Then
		If UCase(WScript.Arguments.Item(2)) = "-FLAGS" Then
			Run WScript.Arguments.Item(1), WScript.Arguments.Item(3)
		End If
	Else
		Run WScript.Arguments.Item(1), ""
	End If
Case "Exec"
	If WScript.Arguments.Count > 2 Then
		If UCase(WScript.Arguments.Item(2)) = "-FLAGS" Then
			Run WScript.Arguments.Item(1), WScript.Arguments.Item(3)
		End If
	Else
		Run WScript.Arguments.Item(1), ""
	End If
Case "MapNetworkDrive"
	MapNetworkDrive WScript.Arguments.Item(1), WScript.Arguments.Item(2), WScript.Arguments.Item(3)
Case "AddNetworkPrinter"
	AddNetworkPrinter WScript.Arguments.Item(1)
Case "SetNetworkPrinterAsDefault"
	SetNetworkPrinterAsDefault WScript.Arguments.Item(1)
Case "CreateDesktopShortcut"
	CreateDesktopShortcut WScript.Arguments.Item(1), WScript.Arguments.Item(2), WScript.Arguments.Item(3)
Case "RedirectShellFolder"
	RedirectShellFolder WScript.Arguments.Item(1), WScript.Arguments.Item(2)
End Select
'Return Successful Exit Code
ExecuteProcedure = 0
End Function

'endregion

'region Procedures
'region File And Folder Operations

'	Create A File
Function CreateFile(filePath, stringToWrite, overwriteBool)
Dim FSO, File
'expand any Environment Variables in filePath string
filePath = EnvVarsInterpreter(filePath)

Set FSO = CreateObject("Scripting.FileSystemObject")
Set File = FSO.CreateTextFile(filePath, overwriteBool)
File.Write(stringToWrite)
File.Close
'Return Successful Exit Code
CreateFile = 0
End Function

'	Create/Append A File
Function AppendFile(filePath, stringToAppend)
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Dim FSO, File
'expand any Environment Variables in filePath string
filePath = EnvVarsInterpreter(filePath)

Set FSO = CreateObject("Scripting.FileSystemObject")
Set File = FSO.OpenTextFile(filePath, ForAppending, True)
File.WriteLine(stringToAppend)
File.Close
'Return Successful Exit Code
AppendFile = 0
End Function

'	Delete Files
Function DeleteFiles(filePath)
Dim FSO
'expand any Environment Variables in filePath string
filePath = EnvVarsInterpreter(filePath)

Set FSO = CreateObject("Scripting.FileSystemObject")
FSO.DeleteFile filePath, True
'Return Successful Exit Code
DeleteFile = 0
End Function
	
'	Create A Folder
Function CreateFolder(targetFolderPath)
Dim FSO
'expand any Environment Variables in targetFolderPath string
targetFolderPath = EnvVarsInterpreter(targetFolderPath)

Set FSO = CreateObject("Scripting.FileSystemObject")
If Not FSO.FolderExists(targetFolderPath) Then 
   FSO.CreateFolder(targetFolderPath)
End If
'Return Successful Exit Code
CreateFolder = 0
End Function

'	Delete Folder
Function DeleteFolder(targetFolderPath)
Dim FSO
'expand any Environment Variables in targetFolderPath string
targetFolderPath = EnvVarsInterpreter(targetFolderPath)

Set FSO = CreateObject("Scripting.FileSystemObject")
If FSO.FolderExists(targetFolderPath) Then
	FSO.DeleteFolder targetFolderPath, True
End If
'Return Successful Exit Code
DeleteFolder = 0
End Function

'	Copy Files
Function CopyFiles(sourceFilePath, targetFilePath, overwriteBool)
Dim FSO
'expand any Environment Variables in sourceFilePath string
sourceFilePath = EnvVarsInterpreter(sourceFilePath)
'expand any Environment Variables in targetFilePath string
targetFilePath = EnvVarsInterpreter(targetFilePath)

Set FSO = CreateObject("Scripting.FileSystemObject")
FSO.CopyFile sourceFilePath, targetFilePath, overwriteBool
'Return Successful Exit Code
CopyFiles = 0
End Function

'	Move Files
Function MoveFiles(sourceFilePath, targetFilePath)
Dim FSO
'expand any Environment Variables in sourceFilePath string
sourceFilePath = EnvVarsInterpreter(sourceFilePath)
'expand any Environment Variables in targetFilePath string
targetFilePath = EnvVarsInterpreter(targetFilePath)

Set FSO = CreateObject("Scripting.FileSystemObject")
FSO.MoveFile sourceFilePath, targetFilePath
'Return Successful Exit Code
MoveFiles = 0
End Function

'	Copy Folder
Function CopyFolder(sourceFolderPath, targetFolderPath, overwriteBool)
Dim FSO
'expand any Environment Variables in sourceFolderPath string
sourceFolderPath = EnvVarsInterpreter(sourceFolderPath)
'expand any Environment Variables in targetFolderPath string
targetFolderPath = EnvVarsInterpreter(targetFolderPath)

Set FSO = CreateObject("Scripting.FileSystemObject")
FSO.CopyFolder sourceFolderPath, targetFolderPath, overwriteBool
'Return Successful Exit code
CopyFolder = 0
End Function

'	Move Folder
Function MoveFolder(sourceFolderPath, targetFolderPath)
Dim FSO
'expand any Environment Variables in sourceFolderPath string
sourceFolderPath = EnvVarsInterpreter(sourceFolderPath)
'expand any Environment Variables in targetFolderPath string
targetFolderPath = EnvVarsInterpreter(targetFolderPath)

Set FSO = CreateObject("Scripting.FileSystemObject")
FSO.MoveFolder sourceFolderPath, targetFolderPath
'Return Successful Exit Code
MoveFolder = 0
End Function

'	Rename A Drive
Function RenameDrive(driveSpec, driveName)
Dim Shell
Set Shell = CreateObject("Shell.Application")
Shell.NameSpace(driveSpec).Self.Name = driveName
'Return Successful Exit Code
RenameDrive = 0
End Function

'	Create A Shortcut
Function CreateShortcut(shortcutName, shortcutFilePath, targetPath)
Dim Shell, Link
'expand any Environment Variables in shortcutFilePath string
shortcutFilePath = EnvVarsInterpreter(shortcutFilePath)

Set Shell = CreateObject("WScript.Shell")
Set link = Shell.CreateShortcut(shortcutFilePath)
With link
.Arguments = ""
.Description = shortcutName
.HotKey = ""
'.IconLocation = targetPath & ", 1"
.TargetPath = targetPath
.WindowStyle = 1
.WorkingDirectory = ""
.Save
End With
'Return Successful Exit Code
CreateShortcut = 0
End Function

'endregion

'region Application Operations

'	Launch a Program
Function Run(commandString, flags)

'use a regex pattern to match a built-in flag "windowstyle="
Dim windowstyleRegEx, windowstyle, match
Set windowstyleRegEx = New RegExp
With windowstyleRegEx
.Pattern = "WINDOWSTYLE=[^ \t\r\n\cK\cL""]*"
.IgnoreCase = True
End With

If windowstyleRegEx.Test(flags) Then
	Set match = windowstyleRegEx.Execute(flags)
	windowstyle = UCase(match(0))
	windowstyle = Replace(windowstyle, "WINDOWSTYLE=", "")

	flags = windowstyleRegEx.Replace(flags, "")
	
	'clean up extra spaces left in flags string
	Dim stringLengthCounter
	stringLengthCounter = Len(flags)
	For i = stringLengthCounter To 2 Step -1
   	spaceCharString = Space(i)
   	flags = Replace(flags, spaceCharString, " ")
	Next

End If

'valid values for windowstyle flag (hidden, minimized, maximized)
Select Case windowStyle
Case "HIDDEN"
	windowStyle = 0
Case "MINIMIZED"
	windowStyle = 2
Case "MAXIMIZED"
	windowStyle = 3
Case Else
	windowStyle = 1
End Select

'DEBUG CODE
'WScript.Echo """" & commandString & """" & " " & flags

'execute requested command with regard to optional flags
Dim WshShell, ExitCode
Set WshShell = CreateObject("WScript.Shell")
ExitCode = WshShell.Run("""" & commandString & """" & " " & flags, windowStyle, False)
'Return Successful Exit Code
Run = ExitCode
End Function

'	Execute A Program
Function Exec(commandString, flags)

'execute requested command with regard to optional flags
'DEBUG CODE
Call AppendFile("\\" & SCRIPTSERVER & "\SYSVOL\" & FQDOMAINNAME & "\scripts\scriptdata\debugResults.txt", "Entering Exec Function")

Dim WshShell, ExecObject, ExitCode, ErrorDesc
Set WshShell = CreateObject("WScript.Shell")
Set ExecObject = WshShell.Exec("""" & commandString & """" & " " & flags)

If Err.Number <> 0 Then 'There was a problem running the command
	ExitCode = Err.Number
	ErrorDesc = "Failed to run command " & commandString & VbCrLf & Err.Description
	Call AppendFile("\\" & SCRIPTSERVER & "\SYSVOL\" & FQDOMAINNAME & "\scripts\scriptdata\execResults.txt", ErrorDesc)
	Exit Function
End If
Do Until ExecObject.Status = 1 'wait until command has completed
	WScript.Sleep(100)
	'DEBUG CODE
Call AppendFile("\\" & SCRIPTSERVER & "\SYSVOL\" & FQDOMAINNAME & "\scripts\scriptdata\debugResults.txt", "Waiting for completion")

Loop
ExitCode = ExecObject.ExitCode
If ExitCode <> 0 Then 'record error info if command did not complete successfully
	ErrorDesc = "Error Code " & ExitCode & ": " & ExecObject.StdErr
	Call AppendFile("\\" & SCRIPTSERVER & "\SYSVOL\" & FQDOMAINNAME & "\scripts\scriptdata\execResults.txt", ErrorDesc)
Else
	Call AppendFile("\\" & SCRIPTSERVER & "\SYSVOL\" & FQDOMAINNAME & "\scripts\scriptdata\execResults.txt", "Command Completed Successfully")
End If
'Return ExitCode
Exec = ExitCode
End Function

'endregion

'region Network Resource Setup Operations

'	Map A Network Drive
Function MapNetworkDrive(unusedDriveLetter, validUNCPath, persistentBool)
Dim FSO, net
Set FSO = CreateObject("Scripting.FileSystemObject")
If Not FSO.DriveExists(unusedDriveLetter) Then
	Set net = CreateObject("WScript.Network")
	net.MapNetworkDrive unusedDriveLetter, validUNCPath, persistentBool
End If
'Return Successful Exit Code
MapNetworkDrive = 0
End Function

'	Add A Network Printer
Function AddNetworkPrinter(networkPrinterUNCPath)
Dim net
Set net = CreateObject("WScript.Network") 
net.AddWindowsPrinterConnection networkPrinterUNCPath
'Return Successful Exit Code
AddNetworkPrinter = 0
End Function

'	Set A Network Printer As Default Printer
Function SetNetworkPrinterAsDefault(networkPrinterUNCPath)
Dim net
Set net = CreateObject("WScript.Network")
net.SetDefaultPrinter networkPrinterUNCPath
'Return Successful Exit Code
SetNetworkPrinterAsDefault = 0
End Function

'endregion

'region User Profile Operations

'	Create A Desktop Shortcut
Function CreateDesktopShortcut(shortcutName, shortcutFileName, targetPath)
Dim Shell, DesktopPath, Link

Set Shell = CreateObject("WScript.Shell")
DesktopPath = Shell.SpecialFolders("Desktop")
Set link = Shell.CreateShortcut(DesktopPath & "\" & shortcutFileName)
With link
.Arguments = ""
.Description = shortcutName
.HotKey = ""
'.IconLocation = targetPath & ", 1"
.TargetPath = targetPath
.WindowStyle = 1
.WorkingDirectory = ""
.Save
End With
'Return Successful Exit Code
CreateDesktopShortcut = 0
End Function



'	Redirect A Shell Folder
Function RedirectShellFolder(nameOfShellFolder, targetFolderPath)
	Dim Shell, regValuePath, regEntryType
	
	Set Shell = CreateObject("WScript.Shell")
	regValuePath = "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\"
	regEntryType = "REG_EXPAND_SZ"
	'determine which folder to redirect
	Select Case nameOfShellFolder
		Case "Desktop"
			regValuePath = regValuePath & "Desktop"
		Case "My Documents"
			regValuePath = regValuePath & "Personal"
		Case "Documents"
			regValuePath = regValuePath & "Personal"
		Case "My Music"
			regValuePath = regValuePath & "My Music"
		Case "Music"
			regValuePath = regValuePath & "My Music"
		Case "My Pictures"
			regValuePath = regValuePath & "My Pictures"
		Case "Pictures"
			regValuePath = regValuePath & "My Pictures"
		Case "My Videos"
			regValuePath = regValuePath & "My Videos"
		Case "Videos"
			regValuePath = regValuePath & "My Videos"
		Case "My Favorites"
			regValuePath = regValuePath & "Favorites"	
		Case "Favorites"
			regValuePath = regValuePath & "Favorites"
	End Select
	
	'if targetPath is not the present path, perform redirection
	Dim currentFolderPath
	currentFolderPath = Shell.RegRead(regValuePath)
	If Not IsNull(currentFolderPath) And Not currentFolderPath = targetFolderPath Then
		'if targetFolderPath does not already exist, create it
		CreateFolder targetFolderPath
		Shell.RegWrite regValuePath, targetFolderPath, regEntryType
	End If
	
End Function

'endregion

'endregion

'region Dictionary containing all the names & required number of arguments of the Procedures contained within this script
Dim procedureDictionary
Set procedureDictionary = CreateObject("Scripting.Dictionary")
procedureDictionary.Add "CreateFile", 3
procedureDictionary.Add "AppendFile", 2
procedureDictionary.Add "DeleteFiles", 1
procedureDictionary.Add "CreateFolder", 1
procedureDictionary.Add "DeleteFolder", 1
procedureDictionary.Add "CopyFiles", 3
procedureDictionary.Add "MoveFiles", 2
procedureDictionary.Add "CopyFolder", 3
procedureDictionary.Add "MoveFolder", 2
procedureDictionary.Add "RenameDrive", 2
procedureDictionary.Add "CreateShortcut", 3
procedureDictionary.Add "Run", 1
procedureDictionary.Add "Exec", 1
procedureDictionary.Add "MapNetworkDrive", 3
procedureDictionary.Add "AddNetworkPrinter", 1
procedureDictionary.Add "SetNetworkPrinterAsDefault", 1
procedureDictionary.Add "CreateDesktopShortcut", 3
procedureDictionary.Add "RedirectShellFolder", 2

'endregion

'region Determine if correct number of arguments were entered for a named Procedure and execute it

'Find number of arguments entered
If WScript.Arguments.Count >1 Then
	Dim procedureName, argc, argv
	procedureName = WScript.Arguments.Item(0)
	argc = WScript.Arguments.Count - 1 'do not count 1st argument (procedureName)
	Dim i
	For i=1 To WScript.Arguments.Count - 1 'store all argument values except procedureName in argv
		argv = argv & WScript.Arguments.Item(i)
	Next
	
	'Determine if number of required arguments is satisfied
	If procedureDictionary.Exists(procedureName) Then
		If procedureDictionary.Item(procedureName) <= argc Then
			
			'Execute Corresponding Procedure
			ExecuteProcedure procedureName, argc
			
		End If
	End If
	
End If
'endregion
'endregion