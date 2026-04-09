' ==============================================================================
' NOMBRE: FormatearExcel.vbs
' DESCRIPCIÓN: Aplica auto-ajuste y negritas a la primera fila de un Excel.
' ARGUMENTOS: 1. Ruta Excel, 2. Hoja (opcional)
' ==============================================================================
Option Explicit
Dim xlApp, xlBook, xlSheet, src, sheetName
If WScript.Arguments.Count < 1 Then WScript.Quit(1)
src = WScript.Arguments(0)
If WScript.Arguments.Count > 1 Then sheetName = WScript.Arguments(1) Else sheetName = 1
Set xlApp = CreateObject("Excel.Application")
Set xlBook = xlApp.Workbooks.Open(src)
Set xlSheet = xlBook.Sheets(sheetName)
xlSheet.Rows(1).Font.Bold = True
xlSheet.UsedRange.Columns.AutoFit
xlBook.Save
xlBook.Close: xlApp.Quit
WScript.Echo "ÉXITO: Formato aplicado."
