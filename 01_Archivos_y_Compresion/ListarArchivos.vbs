' ==============================================================================
' NOMBRE: ListarArchivos.vbs
' DESCRIPCIÓN: Genera una lista de nombres de archivos en una carpeta.
' ARGUMENTOS: 1. Ruta carpeta, 2. Extensión (opcional, ej: .xlsx)
' SALIDA: Lista de nombres separados por nueva línea.
' ==============================================================================
Option Explicit
Dim fso, folder, file, folderPath, ext
If WScript.Arguments.Count < 1 Then WScript.Quit(1)
folderPath = WScript.Arguments(0)
ext = ""
If WScript.Arguments.Count > 1 Then ext = LCase(WScript.Arguments(1))
Set fso = CreateObject("Scripting.FileSystemObject")
If fso.FolderExists(folderPath) Then
    Set folder = fso.GetFolder(folderPath)
    For Each file In folder.Files
        If ext = "" Or LCase(Right(file.Name, Len(ext))) = ext Then
            WScript.Echo file.Path
        End If
    Next
End If
