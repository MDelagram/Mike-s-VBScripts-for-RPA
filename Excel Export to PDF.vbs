'To run this script use:
'wscript|cscript.exe "<FilePathToScript>\Export Word to PDF.vbs" "<filePathToDoc>.docx"

' Check if the command-line arguments are provided
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Please provide the path to the Word document as an argument."
    WScript.Quit 1
End If

'MsgBox "Argument text: " &WScript.Arguments(0)

' Get the path to the Word document from the command-line argument
Dim docPath
docPath = WScript.Arguments(0)

Dim objWord, objDoc

' Create an instance of Word application
Set objWord = CreateObject("Word.Application")

' Hide Word application window
objWord.Visible = False

' Open the Word document
Set objDoc = objWord.Documents.Open(docPath)

' Output PDF path
Dim pdfPath
pdfPath = Replace(docPath,".docx","")

' Export the document as PDF
objDoc.ExportAsFixedFormat pdfPath & ".pdf", 17 ' 17 represents the PDF format


' Close the Word document
objDoc.Close False

' Quit the Word application
objWord.Quit

' Release the objects from memory
Set objDoc = Nothing
Set objWord = Nothing
