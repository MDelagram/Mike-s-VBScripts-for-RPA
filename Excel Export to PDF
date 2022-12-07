'To Open Excel:
Set objExcel = CreateObject("Excel.Application")
objExcel.application.visible = True
objExcel.application.displayalerts = False


'To Attach to an already open Excel:
'Set objExcel = GetObject("C:\Users\myUsername\Desktop\MyExcel.xlsx").Application

Set wkbk = objExcel.Workbooks.Open("C:\Users\MyUsername\Desktop\MyFile.xlsx")
'Set sht = wkbk.Sheets(1)

const xlTypePDF	= 0
const xlQualityStandard = 0
'ExportAsFixedFormat Arguments: 1Type, 2FileName, 3Quality (opt), 4IncludeDocProperties (opt), 5IgnorePrintAreas (opt), 6From (opt), 
'	7To (opt), 8OpenAfterPublish (opt)
objExcel.ActiveWorkbook.ExportAsFixedFormat xlTypePDF,"C:\Users\MyUsername\Desktop\MyFile.pdf",xlQualityStandard,True,False,,,False

'objExcel.ActiveWorkbook.ExportAsFixedFormat xlTypePDF, "C:\Users\mdelagrammatikas\Desktop\MyFile.pdf",xlQualityStandard,1,0,,,0

' Close without saving
objExcel.Application.Quit
objExcel.Quit
