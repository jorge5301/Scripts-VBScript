' ==============================================================================
' NOMBRE: FormatoMesCapitalizado.vbs
' DESCRIPCIÓN: Toma una fecha y devuelve el nombre del mes con la primera en Mayúscula.
' ARGUMENTOS: 1. Fecha (ej: 09/04/2026 o 2026-04-09)
' SALIDA: Nombre del mes (ej: Abril)
' ==============================================================================
Option Explicit

Dim inputDate, monthLabel

If WScript.Arguments.Count < 1 Then
    ' Si no hay argumentos, usar la fecha actual como fallback
    inputDate = Now
Else
    On Error Resume Next
    inputDate = CDate(WScript.Arguments(0))
    If Err.Number <> 0 Then
        WScript.Echo "ERROR: Formato de fecha no válido."
        WScript.Quit(1)
    End If
    On Error GoTo 0
End If

' MonthName devuelve el nombre del mes según el locale del sistema
monthLabel = MonthName(Month(inputDate))

' Asegurar Primera Letra en Mayúscula
If Len(monthLabel) > 0 Then
    monthLabel = UCase(Left(monthLabel, 1)) & LCase(Mid(monthLabel, 2))
End If

WScript.StdOut.Write monthLabel
