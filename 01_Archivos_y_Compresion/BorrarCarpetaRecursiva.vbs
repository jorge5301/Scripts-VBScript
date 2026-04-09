' ==============================================================================
' NOMBRE: BorrarCarpetaRecursiva.vbs
' DESCRIPCIÓN: Elimina una carpeta y todo su contenido.
' ARGUMENTOS: 1. Ruta carpeta
' ==============================================================================
Option Explicit
Dim fso, folderPath
If WScript.Arguments.Count < 1 Then WScript.Quit(1)
folderPath = WScript.Arguments(0)
Set fso = CreateObject("Scripting.FileSystemObject")
If fso.FolderExists(folderPath) Then
    fso.DeleteFolder folderPath, True
    WScript.Echo "ÉXITO: Carpeta eliminada."
Else
    WScript.Echo "INFO: La carpeta no existe."
End If
