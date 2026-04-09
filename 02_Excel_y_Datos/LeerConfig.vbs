' ==============================================================================
' NOMBRE: LeerConfig.vbs
' DESCRIPCIÓN: Lee un valor de un archivo de configuración (Formato Clave=Valor).
' ARGUMENTOS: 1. Ruta archivo, 2. Clave a buscar
' ==============================================================================
Option Explicit
Dim fso, file, line, key, value, found
If WScript.Arguments.Count < 2 Then WScript.Quit(1)
key = LCase(WScript.Arguments(1))
Set fso = CreateObject("Scripting.FileSystemObject")
If Not fso.FileExists(WScript.Arguments(0)) Then WScript.Quit(1)
Set file = fso.OpenTextFile(WScript.Arguments(0), 1)
found = False
Do Until file.AtEndOfStream
    line = Trim(file.ReadLine)
    If InStr(line, "=") > 0 Then
        If LCase(Trim(Left(line, InStr(line, "=") - 1))) = key Then
            value = Trim(Mid(line, InStr(line, "=") + 1))
            WScript.StdOut.Write value
            found = True
            Exit Do
        End If
    End If
Loop
If Not found Then WScript.Echo "CLAVE_NO_ENCONTRADA"
file.Close
