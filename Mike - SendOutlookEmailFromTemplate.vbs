Set oOutlook = CreateObject("Outlook.Application")
Set oMail = oOutlook.CreateItem(0)
TemplatePath = "C:\Users\MyUsername\Desktop\MyOutlookEmailTemplate.oft"

Set oMail = oOutlook.CreateItemFromTemplate(TemplatePath)

oMail.Subject = "Subject"
oMail.To = "myEmailRecepient@myDomain.com"
oMail.CC = ";" 'clear cc list
oMail.BCC = ";"

'oMail.Body = "" 'empty to clear any email singature(s)

'First Display message for signature to be added. Then, use the line below which concatenates the existing bodty to the new body
'oMail.Body = "Dear Sender " & vbrlf & olMail.body

'oMail.HTMLBody = "<html><head><META HTTP-EQUIV=""Content-Type"" CONTENT=""text/html;charset=Windows-1252""><title>Hello</title></head><body><b>Hello World</b></body></html>"

'Add email signature to tml body
'oMail.HTMLBody = "Dear Someone<br>" & "<br>" & oMail.HTMLBody

'oMail.Attachments.Add "C:\Users\MyUsername\Desktop\MyAttachmentFile.extension"
'oMail.SaveAs "C:\Users\MyUsername\Desktop", olSaveAsType.olTemplate


'oMail.Send '(Enable to actually send the email)

oMail.Display

Set oMail = Nothing
Set oOutlook = Nothing