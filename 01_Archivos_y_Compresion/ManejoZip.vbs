' ==============================================================================
' NOMBRE: ManejoZip.vbs
' DESCRIPCIÓN: Comprime o descomprime archivos usando Shell.Application.
' ARGUMENTOS: 1. Acción (zip/unzip), 2. Origen, 3. Destino
' EJEMPLO: cscript ManejoZip.vbs zip "C:\Logs" "C:\Backup\Logs.zip"
' ==============================================================================
Option Explicit
Dim action, src, dst, fso, shell, i
If WScript.Arguments.Count < 3 Then WScript.Quit(1)
action = LCase(WScript.Arguments(0))
src = WScript.Arguments(1)
dst = WScript.Arguments(2)
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("Shell.Application")

If action = "zip" Then
    ' Crear archivo zip vacío con header
    Dim zipFile
    Set zipFile = fso.CreateTextFile(dst, True)
    zipFile.Write "PK" & Chr(5) & Chr(6) & String(18, Chr(0))
    zipFile.Close
    ' Copiar items al zip
    shell.NameSpace(dst).CopyHere shell.NameSpace(src).Items
    ' Esperar a que termine (async)
    Do Until shell.NameSpace(dst).Items.Count = shell.NameSpace(src).Items.Count
        WScript.Sleep 500
    Loop
ElseIf action = "unzip" Then
    If Not fso.FolderExists(dst) Then fso.CreateFolder(dst)
    shell.NameSpace(dst).CopyHere shell.NameSpace(src).Items, 16
    WScript.Sleep 2000 ' Pausa de seguridad para descompresión
End If
WScript.Echo "ÉXITO: Operación " & action & " completada."
