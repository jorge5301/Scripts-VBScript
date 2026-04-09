' ==============================================================================
' NOMBRE: GestionProcesos.vbs
' DESCRIPCIÓN: Verifica si un proceso existe o lo termina forzosamente.
' ARGUMENTOS: 1. Acción (check/kill), 2. Nombre del proceso (ej. excel.exe)
' EJEMPLO: cscript GestionProcesos.vbs kill excel.exe
' ==============================================================================

Option Explicit

Main()

Sub Main()
    If WScript.Arguments.Count < 2 Then
        WScript.Echo "ERROR: Se requieren argumentos: [action] [process_name]"
        WScript.Quit(1)
    End If

    Dim action, processName
    action = LCase(WScript.Arguments(0))
    processName = LCase(WScript.Arguments(1))

    Select Case action
        Case "check"
            If IsProcessRunning(processName) Then
                WScript.Echo "TRUE"
            Else
                WScript.Echo "FALSE"
            End If
        Case "kill"
            KillProcess processName
        Case Else
            WScript.Echo "ERROR: Acción no válida. Use 'check' o 'kill'."
            WScript.Quit(1)
    End Select
End Sub

Function IsProcessRunning(processName)
    Dim objWMIService, colProcesses
    Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
    Set colProcesses = objWMIService.ExecQuery("Select * from Win32_Process Where Name = '" & processName & "'")
    IsProcessRunning = (colProcesses.Count > 0)
End Function

Sub KillProcess(processName)
    Dim objWMIService, colProcesses, objProcess
    Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
    Set colProcesses = objWMIService.ExecQuery("Select * from Win32_Process Where Name = '" & processName & "'")
    
    Dim count
    count = 0
    For Each objProcess in colProcesses
        objProcess.Terminate()
        count = count + 1
    Next
    
    If count > 0 Then
        WScript.Echo "ÉXITO: Se terminaron " & count & " instancias de " & processName
    Else
        WScript.Echo "INFO: No se encontraron instancias de " & processName & " para terminar."
    End If
End Sub
