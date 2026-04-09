' ==============================================================================
' NOMBRE: BuscarReemplazar.vbs
' DESCRIPCIÓN: Busca y reemplaza una cadena de texto dentro de un archivo.
' ARGUMENTOS: 1. Ruta archivo, 2. Texto a buscar, 3. Texto nuevo
' ==============================================================================
Option Explicit
Dim fso, filePath, searchTxt, replaceTxt, content
If WScript.Arguments.Count < 3 Then WScript.Quit(1)
filePath = WScript.Arguments(0)
searchTxt = WScript.Arguments(1)
replaceTxt = WScript.Arguments(2)
Set fso = CreateObject("Scripting.FileSystemObject")
If Not fso.FileExists(filePath) Then WScript.Quit(1)
content = fso.OpenTextFile(filePath, 1).ReadAll
content = Replace(content, searchTxt, replaceTxt)
fso.OpenTextFile(filePath, 2).Write content
WScript.Echo "ÉXITO: Reemplazo completado en " & filePath
