' ==============================================================================
' NOMBRE: UnirExcel.vbs
' DESCRIPCIÓN: Une varios archivos Excel en uno solo (en hojas distintas).
' ARGUMENTOS: 1. Ruta Destino, 2...N. Rutas de Excels origen
' ==============================================================================
Option Explicit
Dim xlApp, xlDest, xlSrc, i, dstPath
If WScript.Arguments.Count < 2 Then WScript.Quit(1)
dstPath = WScript.Arguments(0)
Set xlApp = CreateObject("Excel.Application")
xlApp.DisplayAlerts = False
Set xlDest = xlApp.Workbooks.Add
For i = 1 To WScript.Arguments.Count - 1
    Set xlSrc = xlApp.Workbooks.Open(WScript.Arguments(i))
    xlSrc.Sheets(1).Copy xlDest.Sheets(xlDest.Sheets.Count)
    xlSrc.Close False
Next
xlDest.SaveAs dstPath
xlDest.Close: xlApp.Quit
WScript.Echo "ÉXITO: Excels unidos en " & dstPath
