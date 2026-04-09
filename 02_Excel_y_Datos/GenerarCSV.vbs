' ==============================================================================
' NOMBRE: GenerarCSV.vbs
' DESCRIPCIÓN: Toma argumentos y los escribe como una línea CSV.
' ARGUMENTOS: 1. Ruta CSV, 2...N. Valores
' ==============================================================================
Option Explicit
Dim fso, filePath, i, line, file
If WScript.Arguments.Count < 2 Then WScript.Quit(1)
filePath = WScript.Arguments(0)
line = ""
For i = 1 To WScript.Arguments.Count - 1
    line = line & """" & Replace(WScript.Arguments(i), """", """""") & """"
    If i < WScript.Arguments.Count - 1 Then line = line & ","
Next
Set fso = CreateObject("Scripting.FileSystemObject")
Set file = fso.OpenTextFile(filePath, 8, True) ' 8 = Append
file.WriteLine line
file.Close
WScript.Echo "ÉXITO: Línea añadida a " & filePath
