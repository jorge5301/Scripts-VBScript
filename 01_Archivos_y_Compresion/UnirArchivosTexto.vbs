' ==============================================================================
' NOMBRE: UnirArchivosTexto.vbs
' DESCRIPCIÓN: Une múltiples archivos de texto en uno solo.
' ARGUMENTOS: 1. Ruta Destino, 2...N. Rutas Origen
' ==============================================================================
Option Explicit
Dim fso, dst, i, src, fileDst, fileSrc
If WScript.Arguments.Count < 2 Then WScript.Quit(1)
Set fso = CreateObject("Scripting.FileSystemObject")
Set fileDst = fso.OpenTextFile(WScript.Arguments(0), 8, True)
For i = 1 To WScript.Arguments.Count - 1
    If fso.FileExists(WScript.Arguments(i)) Then
        Set fileSrc = fso.OpenTextFile(WScript.Arguments(i), 1)
        fileDst.WriteLine "--- INICIO ARCHIVO: " & WScript.Arguments(i) & " ---"
        fileDst.Write fileSrc.ReadAll
        fileDst.WriteLine vbCrLf & "--- FIN ARCHIVO ---"
        fileSrc.Close
    End If
Next
fileDst.Close
WScript.Echo "Archivos unidos."
