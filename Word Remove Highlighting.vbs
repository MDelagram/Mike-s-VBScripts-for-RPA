' Run the script with input argument the path to the Word file
' Example:
' "Word Remove Highlighting.vbs" "path_to_Word_document.docx"

Option Explicit

Dim wordApp, doc, range

' Check if the command-line arguments are provided
If WScript.Arguments.Count < 1 Then
    WScript.Echo "Please provide the path to the Word document as an argument."
    WScript.Quit 1
End If

' Get the path to the Word document from the command-line argument
Dim docPath
docPath = WScript.Arguments(0)

' Create an instance of Word application
Set wordApp = CreateObject("Word.Application")

' Open the document
Set doc = wordApp.Documents.Open(docPath)

' Get the entire document range
Set range = doc.Range

' Remove highlighting from the text
range.HighlightColorIndex = 0 ' wdNoHighlight = 0

' Save and close the document
doc.Save
doc.Close

' Quit Word application
wordApp.Quit

' Clean up objects
Set range = Nothing
Set doc = Nothing
Set wordApp = Nothing