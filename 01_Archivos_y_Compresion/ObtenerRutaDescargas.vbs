' ==============================================================================
' NOMBRE: ObtenerRutaDescargas.vbs
' DESCRIPCIÓN: Obtiene la ruta real de la carpeta de Descargas del usuario actual.
' SALIDA: Muestra la ruta en consola (StdOut).
' ==============================================================================

Option Explicit

Main()

Sub Main()
    Dim objShell, objFolder, downloadsPath
    
    On Error Resume Next
    ' Namespace 16 corresponde a la carpeta de Descargas (FDLID_Downloads)
    Set objShell = CreateObject("Shell.Application")
    Set objFolder = objShell.Namespace(16)
    
    If Not objFolder Is Nothing Then
        downloadsPath = objFolder.Self.Path
        ' Asegurar que termina en backslash
        If Right(downloadsPath, 1) <> "\" Then downloadsPath = downloadsPath & "\"
        WScript.StdOut.Write downloadsPath
    Else
        ' Fallback si Shell.Application no funciona por alguna razón
        Dim wshShell
        Set wshShell = CreateObject("WScript.Shell")
        downloadsPath = wshShell.ExpandEnvironmentStrings("%USERPROFILE%") & "\Downloads\"
        WScript.StdOut.Write downloadsPath
    End If
    
    Set objFolder = Nothing
    Set objShell = Nothing
End Sub
