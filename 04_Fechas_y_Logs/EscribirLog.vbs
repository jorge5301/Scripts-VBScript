' ==============================================================================
' NOMBRE: EscribirLog.vbs
' DESCRIPCIÓN: Escribe un mensaje en un archivo log con marca de tiempo.
' ARGUMENTOS: 1. Ruta log, 2. Mensaje, 3. Nivel (INFO/ERROR/WARN)
' ==============================================================================
Option Explicit
Dim fso, logPath, msg, level, file, stamp
If WScript.Arguments.Count < 2 Then WScript.Quit(1)
logPath = WScript.Arguments(0)
msg = WScript.Arguments(1)
If WScript.Arguments.Count > 2 Then level = UCase(WScript.Arguments(2)) Else level = "INFO"
stamp = Year(Now) & "-" & Right("0" & Month(Now), 2) & "-" & Right("0" & Day(Now), 2) & " " & _
        Right("0" & Hour(Now), 2) & ":" & Right("0" & Minute(Now), 2) & ":" & Right("0" & Second(Now), 2)
Set fso = CreateObject("Scripting.FileSystemObject")
Set file = fso.OpenTextFile(logPath, 8, True)
file.WriteLine "[" & stamp & "] [" & level & "] " & msg
file.Close
WScript.Echo "Log escrito."
