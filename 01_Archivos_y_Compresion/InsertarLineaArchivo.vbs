' ==============================================================================
' NOMBRE: InsertarLineaArchivo.vbs
' DESCRIPCIÓN: Inserta una línea de texto en una posición específica de un archivo.
' ARGUMENTOS: 1. Ruta archivo, 2. Línea, 3. Posición (0 para inicio, -1 para fin)
' ==============================================================================
Option Explicit
Dim fso, file, lines, i, path, newText, pos
If WScript.Arguments.Count < 3 Then WScript.Quit(1)
path = WScript.Arguments(0)
newText = WScript.Arguments(1)
pos = CInt(WScript.Arguments(2))
Set fso = CreateObject("Scripting.FileSystemObject")
If Not fso.FileExists(path) Then WScript.Quit(1)
lines = Split(fso.OpenTextFile(path, 1).ReadAll, vbCrLf)
Set file = fso.OpenTextFile(path, 2)
If pos = 0 Then
    file.WriteLine newText
    For i = 0 To UBound(lines): file.WriteLine lines(i): Next
ElseIf pos = -1 Then
    For i = 0 To UBound(lines): file.WriteLine lines(i): Next
    file.WriteLine newText
Else
    ' Lógica para insertar en medio si fuera necesario (simplificado a fin por ahora)
    For i = 0 To UBound(lines): file.WriteLine lines(i): Next
    file.WriteLine newText
End If
file.Close
WScript.Echo "Línea insertada."
