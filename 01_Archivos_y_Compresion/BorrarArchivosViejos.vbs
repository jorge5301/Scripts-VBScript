' ==============================================================================
' NOMBRE: BorrarArchivosViejos.vbs
' DESCRIPCIÓN: Elimina archivos de una carpeta con antigüedad mayor a X días.
' ARGUMENTOS: 1. Ruta carpeta, 2. Días de antigüedad
' EJEMPLO: cscript BorrarArchivosViejos.vbs "C:\Temp" 7
' ==============================================================================
Option Explicit
Dim fso, folder, file, folderPath, daysOld, count
If WScript.Arguments.Count < 2 Then WScript.Quit(1)
folderPath = WScript.Arguments(0)
daysOld = CInt(WScript.Arguments(1))
Set fso = CreateObject("Scripting.FileSystemObject")
If Not fso.FolderExists(folderPath) Then WScript.Quit(1)
Set folder = fso.GetFolder(folderPath)
count = 0
For Each file In folder.Files
    If DateDiff("d", file.DateLastModified, Now) > daysOld Then
        On Error Resume Next
        file.Delete True
        If Err.Number = 0 Then count = count + 1
        On Error GoTo 0
    End If
Next
WScript.Echo "ÉXITO: Se eliminaron " & count & " archivos viejos."
