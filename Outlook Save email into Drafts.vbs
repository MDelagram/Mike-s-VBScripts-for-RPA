Dim olApp
Dim objMailItem

Set olApp = CreateObject("Outlook.Application")
Set objMailItem = olApp.CreateItem(0)
objMailItem.Subject = "Test"
objMailItem.To = "abc@email.com"
objMailItem.Body = "Test Email"
objMailItem.Save

Set olApp = Nothing
Set objMailItem = Nothing
