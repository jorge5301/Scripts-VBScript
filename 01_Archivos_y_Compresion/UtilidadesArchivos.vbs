' ==============================================================================
' NOMBRE: UtilidadesArchivos.vbs
' DESCRIPCIÓN: Acciones misceláneas sobre archivos para bots.
' ARGUMENTOS: 1. Acción (age/move), 2. Ruta Origen, 3. Ruta Destino o Minutos
' EJEMPLO: cscript UtilidadesArchivos.vbs age "C:\log.txt" 10
' EJEMPLO: cscript UtilidadesArchivos.vbs move "C:\tmp.txt" "C:\backup\tmp.txt"
' ==============================================================================

Option Explicit

Main()

Sub Main()
    If WScript.Arguments.Count < 2 Then
        WScript.Echo "ERROR: Argumentos insuficientes."
        WScript.Quit(1)
    End If

    Dim action, path1
    action = LCase(WScript.Arguments(0))
    path1 = WScript.Arguments(1)

    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")

    Select Case action
        Case "age"
            If WScript.Arguments.Count < 3 Then
                WScript.Echo "ERROR: Especifique los minutos para la acción 'age'."
                WScript.Quit(1)
            End If
            Dim mins, fileObj
            mins = CInt(WScript.Arguments(2))
            If fso.FileExists(path1) Then
                Set fileObj = fso.GetFile(path1)
                ' DateDiff en minutos
                If DateDiff("n", fileObj.DateLastModified, Now) > mins Then
                    WScript.Echo "TRUE" ' Es más viejo que X minutos
                Else
                    WScript.Echo "FALSE"
                End If
            Else
                WScript.Echo "ERROR: Archivo no existe."
            End If
            
        Case "move"
            If WScript.Arguments.Count < 3 Then
                WScript.Echo "ERROR: Especifique el destino para la acción 'move'."
                WScript.Quit(1)
            End If
            Dim path2
            path2 = WScript.Arguments(2)
            MoverConReintento path1, path2, 5 ' 5 reintentos por defecto

        Case Else
            WScript.Echo "ERROR: Acción no válida."
    End Select
End Sub

Sub MoverConReintento(origen, destino, reintentos)
    Dim fso, i, exito
    Set fso = CreateObject("Scripting.FileSystemObject")
    exito = False
    
    For i = 1 To reintentos
        On Error Resume Next
        If fso.FileExists(destino) Then fso.DeleteFile destino, True
        fso.MoveFile origen, destino
        If Err.Number = 0 Then
            exito = True
            Exit For
        End If
        WScript.Sleep 2000 ' Esperar 2 segundos antes de reintentar
    Next
    
    If exito Then
        WScript.Echo "ÉXITO: Archivo movido a " & destino
    Else
        WScript.Echo "ERROR: No se pudo mover el archivo después de " & reintentos & " intentos."
        WScript.Quit(1)
    End If
End Sub
