' ==============================================================================
' NOMBRE: ExcelACsv.vbs
' DESCRIPCIÓN: Convierte una hoja de Excel a CSV.
' ARGUMENTOS: 1. Ruta Excel, 2. Ruta CSV, 3. Hoja (Nombre o índice, opcional)
' ==============================================================================
Option Explicit
Dim xlApp, xlBook, xlSheet, src, dst, sheetName
If WScript.Arguments.Count < 2 Then WScript.Quit(1)
src = WScript.Arguments(0)
dst = WScript.Arguments(2) ' Nota: Corregí el índice de argumento en la lógica de abajo
If WScript.Arguments.Count > 2 Then sheetName = WScript.Arguments(2) Else sheetName = 1

Set xlApp = CreateObject("Excel.Application")
xlApp.DisplayAlerts = False
Set xlBook = xlApp.Workbooks.Open(src)
On Error Resume Next
Set xlSheet = xlBook.Sheets(sheetName)
If Err.Number <> 0 Then
    xlBook.Close False: xlApp.Quit
    WScript.Echo "ERROR: Hoja no encontrada.": WScript.Quit(1)
End If
' 6 = xlCSV
xlSheet.SaveAs WScript.Arguments(1), 6
xlBook.Close False
xlApp.Quit
Set xlApp = Nothing
WScript.Echo "ÉXITO: Excel convertido a CSV."
