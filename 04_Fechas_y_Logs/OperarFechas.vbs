' ==============================================================================
' NOMBRE: OperarFechas.vbs
' DESCRIPCIÓN: Suma o resta tiempo a una fecha.
' ARGUMENTOS: 1. FechaBase, 2. Cantidad (ej: 5), 3. Unidad (d/m/yyyy/h/n)
' ==============================================================================
Option Explicit
Dim f, val, unit
If WScript.Arguments.Count < 3 Then WScript.Quit(1)
f = CDate(WScript.Arguments(0))
val = CInt(WScript.Arguments(1))
unit = WScript.Arguments(2)
WScript.StdOut.Write DateAdd(unit, val, f)
