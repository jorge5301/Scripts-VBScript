' ==============================================================================
' NOMBRE: DiferenciaFechas.vbs
' DESCRIPCIÓN: Calcula la diferencia entre dos fechas en la unidad especificada.
' ARGUMENTOS: 1. Fecha1, 2. Fecha2, 3. Unidad (d/h/n/s)
' ==============================================================================
Option Explicit
Dim f1, f2, unit
If WScript.Arguments.Count < 3 Then WScript.Echo "ERROR: Argumentos insuficientes": WScript.Quit(1)
f1 = CDate(WScript.Arguments(0))
f2 = CDate(WScript.Arguments(1))
unit = WScript.Arguments(2)
WScript.StdOut.Write DateDiff(unit, f1, f2)
