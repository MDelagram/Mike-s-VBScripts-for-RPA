' This script sends an Email from Outlook Default Account on behalf of another email account or group email.
' It requires Outlook to be configured to Send Ekail on behalf of another user
' Check out the following link for instructions: https://www.eduhk.hk/ocio/content/faq-how-send-mail-behalf-another-user

'Prepare Outlook Email
Function Prepare_Mail() 
	Set objOutlook = CreateObject("Outlook.Application")
	Set objMail = objOutlook.CreateItem(0)
	
	'On Error Resume Next
	With objMail
		'Set your own email addresses in the 'To' and 'SentOnBehalfOfName' properties
		.to = "receipient_email@domain.com"
		.SentOnBehalfOfName = "Email_Account_To_Send_On_their_behalf"
		'.Cc = ""
		'.Bcc = ""
		.Subject = "VBScript Test Email"
		.Body = "Type Message Body"

		'.HTMLBody = ""
		'.Attachments.Add ActiveWorkbook.FullName
		.display	
		.Send 'sends the email
	End With
	'On Error GoTO 0
	
	Set objMail = Nothing
	Set objOutlook = Nothing
End Function


'Main{
	'Call functions
	Call Prepare_Mail()
'}	
