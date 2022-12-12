'every 2 minutes toggles the Numlock [On/Off] to trick the system into believing that someone is interacting with the UI
'There is also the option to choose using the F13 key instead
'source: https://superuser.com/questions/329758/how-can-i-prevent-a-policy-enforced-screen-lock-in-windows/836346#836346

Dim objResult

Set objShell = WScript.CreateObject("WScript.Shell")    

Do While True
	objResult = objShell.SendKeys("{NUMLOCK}")
	'objResult = objShell.SendKeys("{F13}") 'F13 function key is unused in Windows. It is ment for future use.
	Wscript.Sleep(30) 'for slightly visible numlock changes (enable/disable as needed)
	objResult = objShell.SendKeys("{NUMLOCK}")
	Wscript.Sleep(120000) 'every 2 minutes in MS 2*60*1000
Loop
