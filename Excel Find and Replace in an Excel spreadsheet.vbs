'Find and Replace in an Excel spreadsheet

Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True

Set objWorkbook = objExcel.Workbooks.Open("C:\Users\MyUsername\Desktop\Test.xlsx")

'Select the first (or whatever) worksheet of the Excel file
Set objWorksheet = objWorkbook.Worksheets(1)

'UsedRange property create an instance of the Range object that automatically includes every cell that has data in it:
Set objRange = objWorksheet.UsedRange

'text we want to search, text we want to replace it with (if text does not exist, it doesn't do anything).
'Replace "C:\Test\Image.jpg" with "C:\Backup\Image.jpg"
objRange.Replace "C:\Test\Image.jpg", "C:\Backup\Image.jpg"


