' ==============================================================================
' NOMBRE: AsegurarMinusculas.vbs
' DESCRIPCIÓN: Verifica si Caps Lock está activo y lo desactiva si es necesario.
' REQUISITOS: Requiere Microsoft Word instalado para validación de estado real.
' USO: Ejecutar al inicio de tareas que requieran ingreso de texto específico.
' ==============================================================================

Option Explicit

Main()

Sub Main()
    On Error Resume Next
    Dim capsState
    capsState = GetCapsLockState()
    
    ' Si no se pudo determinar el estado (Null), no hacemos nada por seguridad
    If IsNull(capsState) Then
        WScript.Echo "ADVERTENCIA: No se pudo determinar el estado de Caps Lock (Word no instalado)."
        WScript.Quit(1)
    End If
    
    ' Si está encendido (True), lo apagamos
    If capsState = True Then
        DesactivarCapsLock()
        WScript.Echo "ÉXITO: Caps Lock ha sido desactivado."
    Else
        WScript.Echo "INFO: Caps Lock ya estaba desactivado."
    End If
End Sub

Function GetCapsLockState()
    Dim objWord
    On Error Resume Next
    Set objWord = CreateObject("Word.Application")
    If Err.Number <> 0 Then
        GetCapsLockState = Null
    Else
        GetCapsLockState = objWord.CapsLock
        objWord.Quit
    End If
    Set objWord = Nothing
End Function

Sub DesactivarCapsLock()
    Dim objShell
    Set objShell = CreateObject("WScript.Shell")
    objShell.SendKeys "{CAPSLOCK}"
    Set objShell = Nothing
End Sub
