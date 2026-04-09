' ==============================================================================
' NOMBRE: CrearArchivoVacio.vbs
' DESCRIPCIÓN: Crea un archivo vacío (0 bytes). Útil para control de bots (Flags).
' ARGUMENTOS: 1. Ruta del archivo
' ==============================================================================
Option Explicit
Dim fso, filePath
If WScript.Arguments.Count < 1 Then WScript.Quit(1)
filePath = WScript.Arguments(0)
Set fso = CreateObject("Scripting.FileSystemObject")
fso.CreateTextFile(filePath, True).Close
WScript.Echo "ÉXITO: Archivo flag creado en " & filePath
