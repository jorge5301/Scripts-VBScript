' ==============================================================================
' NOMBRE: ManejoRar.vbs
' DESCRIPCIÓN: Comprime o descomprime archivos .rar (y otros) usando WinRAR o 7-Zip.
' REQUISITOS: Requiere WinRAR o 7-Zip instalado en la máquina del bot.
' ARGUMENTOS: 1. Acción (extract/compress), 2. Ruta Origen, 3. Ruta Destino
' ==============================================================================
Option Explicit

Main()

Sub Main()
    If WScript.Arguments.Count < 3 Then
        WScript.Echo "ERROR: Faltan argumentos. Uso: extract/compress [Origen] [Destino]"
        WScript.Quit(1)
    End If

    Dim action, src, dst, toolExec, shell
    action = LCase(WScript.Arguments(0))
    src = WScript.Arguments(1)
    dst = WScript.Arguments(2)
    
    Set shell = CreateObject("WScript.Shell")
    
    ' Localizar herramienta (WinRAR o 7-Zip)
    toolExec = BuscarHerramienta()
    
    If toolExec = "" Then
        WScript.Echo "ERROR: No se encontró WinRAR o 7-Zip instalado en las rutas por defecto."
        WScript.Quit(1)
    End If

    Dim command
    If action = "extract" Then
        ' x: Extraer con rutas completas
        ' -y: Asumir sí a todo (sobreescribir)
        If InStr(LCase(toolExec), "7z.exe") > 0 Then
            command = """" & toolExec & """ x """ & src & """ -o""" & dst & """ -y"
        Else
            command = """" & toolExec & """ x -ibck """ & src & """ """ & dst & "\"""
        End If
    ElseIf action = "compress" Then
        ' a: Añadir al archivo
        command = """" & toolExec & """ a """ & dst & """ """ & src & """"
    Else
        WScript.Echo "ERROR: Acción no válida (use extract o compress)."
        WScript.Quit(1)
    End If

    WScript.Echo "Ejecutando: " & action
    Dim exitCode
    exitCode = shell.Run(command, 0, True)

    If exitCode = 0 Then
        WScript.Echo "ÉXITO: Operación completada."
    Else
        WScript.Echo "ERROR: El comando falló con el código " & exitCode
    End If
End Sub

Function BuscarHerramienta()
    Dim fso, paths, p
    Set fso = CreateObject("Scripting.FileSystemObject")
    ' Rutas comunes de búsqueda
    paths = Array( _
        "C:\Program Files\7-Zip\7z.exe", _
        "C:\Program Files (x86)\7-Zip\7z.exe", _
        "C:\Program Files\WinRAR\WinRAR.exe", _
        "C:\Program Files (x86)\WinRAR\WinRAR.exe" _
    )
    
    For Each p In paths
        If fso.FileExists(p) Then
            BuscarHerramienta = p
            Exit Function
        End If
    Next
    BuscarHerramienta = ""
End Function
