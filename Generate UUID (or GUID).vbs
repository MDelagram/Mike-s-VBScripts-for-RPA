' This script generates a UUID or 'Universally Unique Identifier'. This is used anywhere a unique identifier is required eg. for component identifiers, database keys, serial numbers, etc.)
' source: https://superuser.com/questions/155740/how-can-i-generate-a-uuid-from-the-command-line-in-windows-xp

set obj = CreateObject("Scriptlet.TypeLib")
MsgBox "UUID generated is: " & obj.GUID

' in case you need to remove curly braces {}
MsgBox Replace(Replace(obj.GUID,"{",""),"}","")