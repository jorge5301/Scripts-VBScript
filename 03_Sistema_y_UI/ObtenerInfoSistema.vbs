' ==============================================================================
' NOMBRE: ObtenerInfoSistema.vbs
' DESCRIPCIÓN: Obtiene información básica del sistema para auditoría.
' SALIDA: JSON-like string con Nombre, Usuario e IP.
' ==============================================================================
Option Explicit
Dim network, computer, user, objWMIService, colIP, objIP, ip
Set network = CreateObject("WScript.Network")
computer = network.ComputerName
user = network.UserName
Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
Set colIP = objWMIService.ExecQuery("Select IPAddress from Win32_NetworkAdapterConfiguration Where IPEnabled = True")
ip = "N/A"
For Each objIP in colIP
    If Not IsNull(objIP.IPAddress) Then ip = objIP.IPAddress(0): Exit For
Next
WScript.Echo "COMPUTER:" & computer & "|USER:" & user & "|IP:" & ip
