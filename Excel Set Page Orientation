'To Open Excel:
Set objExcel = CreateObject("Excel.Application")
objExcel.application.visible = True
objExcel.application.displayalerts = False


'To Attach to an already open Excel:
'Set objExcel = GetObject("C:\Users\MyUsername\Desktop\MyExcel.xlsx").Application

Set wkbk = objExcel.Workbooks.Open("C:\Users\MyUsername\Desktop\MyExcel.xlsx")
'Set sht = wkbk.Sheets(1)

' switch page orientation to Landscape
const xlPageOrientation = 2 'set to 1 for Portrait and 2 for Landscape mode
objExcel.Sheets(1).PageSetup.Orientation = xlPageOrientation

' Close without saving
objExcel.Application.Quit
objExcel.Quit
