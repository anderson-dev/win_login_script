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
'		AppDeployStartupScript.vbs

'### Description ###
'
'		This script is used to install common applications at startup.  It makes use of a central
'		script library (ScriptLibrary.vbs) to abstract the more detailed implementation of common
'		procedures.



'region Script Source Code
Option Explicit
Const SCRIPTSERVER = "ADS1"
Const FQDOMAINNAME = "HOME.LAN"

'region Implements Central Script Library (ScriptLibrary.vbs)
Function ScriptLibrary(procedureNameWithArgsSpaceDelimited)
Dim Shell
Set Shell = CreateObject("WScript.Shell")
Shell.Run("cscript \\" & SCRIPTSERVER & "\SYSVOL\" & FQDOMAINNAME & "\scripts\ScriptLibrary.vbs" & " " & procedureNameWithArgsSpaceDelimited)
End Function

'endregion

Dim scriptPath
scriptPath = Left(WScript.ScriptFullName, Len(WScript.ScriptFullName) - Len(WScript.ScriptName))

'region Determine Optional Arguments Used
	'if no arguments specified
	If WScript.Arguments.Count >= 1 Then
		Select Case WScript.Arguments.Item(0)
			Case "-auto"
				'Auto Deployment Mode
				AutoDeploy()
			Case "-install"
				'Remote install mode
			Case "-add"
				'Add application to deployment list
				If WScript.Arguments.Count = 3 Then
					AddApp WScript.Arguments.Item(1), "", WScript.Arguments.Item(2)
				ElseIf WScript.Arguments.Count > 3 Then
					AddApp WScript.Arguments.Item(1), WScript.Arguments.Item(3), WScript.Arguments.Item(4)
				Else
					WScript.Echo "Incorrect number of arguments given"
				End If
			Case "-delete"
				'Delete application from deployment list
			Case "-remove"
				'Remove application from all machines (also deletes from deployment list)
		End Select
	Else
		'no option was invoked
		WScript.Echo "Please Specify an Option:" & vbCrLf & vbCrLf & _
		"-auto" & vbTab & "Auto Deployment Mode" & vbCrLf & _
		"-install" & vbTab & "Remote one-time Installation Mode" & vbCrLf & _
		"-add" & vbTab & "Add an application to the deployment list" & vbCrLf & _
		"-delete" & vbTab & "Delete an application from the deployment list" & vbCrLf & _
		"-remove" & vbTab & "Remove an application from all machines"
	End If
'endregion


'region Auto Deployment
Sub AutoDeploy()
	'Deploys specified applications to the pc running the startup script
	'(If it Is Not already installed)
	
	'check for existing deployment list
	Dim FSO, File
	Const ForReading = 1
	Set FSO = CreateObject("Scripting.FileSystemObject")
	If FSO.FileExists(scriptPath & "scriptdata\deploymentlist.txt") Then
		'open deploymentlist.txt for reading
		Set File = FSO.OpenTextFile(scriptPath & "scriptdata\deploymentlist.txt", ForReading, False)
		'read through all application deployment entries and install as needed	
		Do Until File.AtEndOfStream
			Dim appInstallPath, flags, postInstallFileCheck
			appInstallPath = File.ReadLine
			flags = File.ReadLine
			postInstallFileCheck = File.ReadLine

			If Not FSO.FileExists(postInstallFileCheck) Then
				Call Exec(appInstallPath, flags, "\\" & SCRIPTSERVER & "\SYSVOL\" & FQDOMAINNAME & "\scripts\scriptdata\execLog.txt")
			End If
		Loop
		File.Close
	Else
				
	End If
	
End Sub
'endregion

'region Add Application to Deployment List
Sub AddApp(appInstallPath, flags, postInstallFileCheck)
	'Adds an application to deploymentlist.txt
	
	'create scriptdata folder if it doesn't already exist
	ScriptLibrary "CreateFolder " & scriptPath & "scriptdata"
	WScript.Sleep(50) 'wait for the folder to be created if it is not there
	'add to deployment list (create it if it does not exist)
	ScriptLibrary "AppendFile " & """" & scriptPath & "scriptdata\deploymentlist.txt" & """" & _
	" " & """" & appInstallPath & """"
	WScript.Sleep(50)
	ScriptLibrary "AppendFile " & """" & scriptPath & "scriptdata\deploymentlist.txt" & """" & _
	" " & """" & flags & """"
	WScript.Sleep(50)
	ScriptLibrary "AppendFile " & """" & scriptPath & "scriptdata\deploymentlist.txt" & """" & _
	" " & """" & postInstallFileCheck & """"

End Sub

'endregion

'region Execute and Log Shell Command
Function Exec(commandString, flags, logFileSpec)

	'Executes a shell command with optional flags
	'an optional logFile can be specified for recording execution errors
	'return value is 0 for successful and non-zero for unsuccessful

	'Declare local variables needed
	Dim shell, execObject, exitCode
	Dim errorDesc
	
	'Execute command
	Set shell = CreateObject("WScript.Shell")
	Set execObject = shell.Exec("""" & commandString & """" & " " & flags)
	
	'If there was a problem running the command exit immediately
	If Err.Number <> 0 Then
		Call AppendLog(logFileSpec, "Failed to run command " & commandString & VbTab & Err.Description)
		Exec = Err.Number
		Exit Function
	End If
	
	'Wait for the command to finish running
	Do Until execObject.Status = 1 'do until finished
		WScript.Sleep(100)
	Loop
	
	'Record command's exit status
	exitCode = execObject.ExitCode
	If exitCode <> 0 Then
		'command had errors
		Call AppendLog(logFileSpec, "Error Code " & exitCode & ": " & execObject.StdErr)
	Else
		'command completed successfully
		Call AppendLog(logFileSpec, "Command Completed Successfully!")
	End If
	
	'Return exitCode
	Exec = exitCode
End Function

'Handle writing to a log file if one is specified
Sub AppendLog(logFileSpec, stringToWrite)
	
	'If logFileSpec is not Empty, try writing log file
	If logFileSpec <> "" Then
		
		'Write log file out
		Const ForReading = 1, ForWriting = 2, ForAppending = 8

		Dim fileObject, textFile

		Set fileObject = CreateObject("Scripting.FileSystemObject")

		Set textFile = fileObject.OpenTextFile(logFileSpec, ForAppending, True)

		textFile.WriteLine(stringToWrite)

		textFile.Close		
		
	End If
		
End Sub

'endregion