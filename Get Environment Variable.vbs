' This script is an example of getting the Environment Variables from the System and using them in VBscript

Set oShell = CreateObject("WScript.Shell")
user = oShell.ExpandEnvironmentStrings("%UserName%")
comp = oShell.ExpandEnvironmentStrings("%ComputerName%")

WScript.Echo "Username: " & user & vbCrLf & "Computer name: " & comp

dim answer
'display information message: parameters 'message', 'buttons type', title
answer = MsgBox("Username: " & user & vbCrLf & "Computer name: " & comp, vbInformation, "Environment Variables")

'Set WshProccessEnv = oShell.Environment("Process")
'Set WshSysEnv = oShell.Environment("System")

'Wscript.Echo WshSysEnv("NUMBER_OF_PROCESSORS") & vbCrLf
'Wscript.Echo WshProccessEnv("Path")

Set oShell = Nothing