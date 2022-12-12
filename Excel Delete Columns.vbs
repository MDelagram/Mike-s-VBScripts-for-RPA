'To Open Excel:
Set objExcel = CreateObject("Excel.Application")
objExcel.application.visible = True
objExcel.application.displayalerts = False


'To Attach to an already open Excel:
'Set objExcel = GetObject("C:\Users\MyUsername\Desktop\Parking Test.xlsx").Application

Set wkbk = objExcel.Workbooks.Open("C:\Users\MyUsername\Desktop\Parking Test.xlsx")
'Set sht = wkbk.Sheets(1)

'Delete Columns
objExcel.ActiveSheet.Range("C:G").Delete
Range("C:C,D:D,E:E,F:F,G:G").Delete


' Close without saving
objExcel.Application.Quit
objExcel.Quit
