' ==============================================================================
' NOMBRE: EsperarVentana.vbs
' DESCRIPCIÓN: Pausa el script hasta que aparezca una ventana con el título dado.
' ARGUMENTOS: 1. Título de la ventana, 2. Tiempo max espera (segundos)
' ==============================================================================
Option Explicit
Dim shell, title, timeout, start, found
If WScript.Arguments.Count < 2 Then WScript.Quit(1)
title = WScript.Arguments(0)
timeout = CInt(WScript.Arguments(1))
Set shell = CreateObject("WScript.Shell")
start = Timer
found = False
Do While Timer < start + timeout
    If shell.AppActivate(title) Then
        found = True
        Exit Do
    End If
    WScript.Sleep 500
Loop
If found Then WScript.Echo "TRUE" Else WScript.Echo "FALSE"
