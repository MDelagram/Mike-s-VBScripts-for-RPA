' This script converts the .rtf file to .docx
' Usage cscript "C:\Users\<username>\Desktop\Convert RTF to Word.vbs" "C:\Users\<username>\Desktop\Test.rtf"

Option Explicit

' Constants for file format
Const wdFormatRTF = 6 ' for .rtf extension
Const wdFormatXMLDocument = 12 ' for .docx extension

' Check for the correct number of command-line arguments
If WScript.Arguments.Count <> 1 Then
    MsgBox "Usage: cscript <path to script.vbs> <path to .rtf file>"
    WScript.Quit 1
End If

' Input and output file paths
Dim inputFile, outputFile

' Input file path (passed as a command-line argument)
inputFile = WScript.Arguments(0)

' Set Output Word file path to be saved in the same folder as the input file.
outputFile = Replace(inputFile,".rtf",".docx")


' Validate the input file extension (must be .rtf)
Dim fso, inputExtension
Set fso = CreateObject("Scripting.FileSystemObject")
inputExtension = LCase(fso.GetExtensionName(inputFile))
If inputExtension <> "rtf" Then
    MsgBox "Invalid input file format. The input file must be an .rtf file.", vbExclamation, "Invalid File Format"
    WScript.Quit 1
End If


' Create Word Application object
Dim objWord
Set objWord = CreateObject("Word.Application")

' Hide Word Application window (optional, uncomment the next line to hide)
'objWord.Visible = False

' Function to wait for file creation
Function WaitForFileCreation(filePath)
    Dim fso, file, maximumRetries, retriesCounter
    Set fso = CreateObject("Scripting.FileSystemObject")
    WaitForFileCreation = False
	
	maximumRetries = 200
	retriesCounter = 0
	

    Do While Not fso.FileExists(filePath)
		' MsgBox "File does not exist yet will retry " &(retriesCounter +1)& " of " &maximumRetries& "."
        WScript.Sleep 100 ' Wait for 100 milliseconds before checking again
		
		' Check if the maximum number of retries has been exceeded. (100 ms delay * MaximumRetries (200) = 20 sec.)
		If retriesCounter > maximumRetries Then
			Exit Function ' Exit the function with WaitForFileCreation = False
		End If
		
		retriesCounter = retriesCounter + 1
    Loop

    WaitForFileCreation = True
End Function


' Open the RTF file
    Dim objDocument
    Set objDocument = objWord.Documents.Open(inputFile)

    ' Save the document in .docx format
    objDocument.SaveAs2 outputFile, wdFormatXMLDocument

    ' Close the document and quit Word
    objDocument.Close
    objWord.Quit

    ' Release objects from memory
    Set objDocument = Nothing
    Set objWord = Nothing

' Wait for the input file to be created
If WaitForFileCreation(outputFile) Then    
    ' Inform the user about the successful conversion
    MsgBox "Conversion completed! The .docx file is saved at " & outputFile, vbInformation, "Conversion Status"
Else
    MsgBox "The output file was not found.", vbExclamation, "File Not Found"
End If
