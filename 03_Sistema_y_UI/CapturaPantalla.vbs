' ==============================================================================
' NOMBRE: CapturaPantalla.vbs
' DESCRIPCIÓN: Toma una captura de pantalla completa y la guarda como PNG.
' ARGUMENTOS: 1. Ruta destino (ej: C:\Evidencia\Error.png)
' ==============================================================================
Option Explicit
Dim shell, filePath, psCommand
If WScript.Arguments.Count < 1 Then WScript.Quit(1)
filePath = WScript.Arguments(0)
Set shell = CreateObject("WScript.Shell")

' Comando PowerShell para capturar pantalla
psCommand = "powershell -Command ""Add-Type -AssemblyName System.Windows.Forms; " & _
            "[System.Windows.Forms.SendKeys]::SendWait('{PRTSC}'); " & _
            "Start-Sleep -m 500; " & _
            "If([System.Windows.Forms.Clipboard]::ContainsImage()){ " & _
            "$image = [System.Windows.Forms.Clipboard]::GetImage(); " & _
            "$image.Save('" & filePath & "', [System.Drawing.Imaging.ImageFormat]::Png); " & _
            "Write-Output 'ÉXITO' } Else { Write-Output 'ERROR' }"""

shell.Run psCommand, 0, True
WScript.Echo "Captura procesada en " & filePath
