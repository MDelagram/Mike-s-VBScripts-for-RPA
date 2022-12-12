' Used sometimes when we need to find to get the path of regardless windows Username.

set WshShell = WScript.CreateObject("WScript.Shell")
strDesktop = WshShell.SpecialFolders("Desktop") 'Has other uses to. Change 'Desktop' to any of the below (not all available depending on Windows version):

'   AllUsersDesktop
'   AllUsersStartMenu
'   AllUsersPrograms
'   AllUsersStartup
'   Desktop
'   Favorites
'   Fonts
'   MyDocuments
'   StartMenu
'   Startup

wscript.echo strDesktop

MsgBox "Special folder " &strDesktop
