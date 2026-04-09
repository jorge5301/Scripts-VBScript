' ==============================================================================
' NOMBRE: SimularTeclas.vbs
' DESCRIPCIÓN: Envía pulsaciones de teclas a la ventana activa.
' ARGUMENTOS: 1. Teclas (ej: ^c para Ctrl+C, %{TAB} para Alt+Tab)
' ==============================================================================
Option Explicit
Dim shell, keys
If WScript.Arguments.Count < 1 Then WScript.Quit(1)
keys = WScript.Arguments(0)
Set shell = CreateObject("WScript.Shell")
shell.SendKeys keys
WScript.Echo "ÉXITO: Teclas enviadas: " & keys
