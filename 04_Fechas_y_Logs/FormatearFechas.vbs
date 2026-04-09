' ==============================================================================
' NOMBRE: FormatearFechas.vbs
' DESCRIPCIÓN: Devuelve la fecha o hora actual en formatos estandarizados.
' ARGUMENTOS: 1. Formato (iso / dd-mm-yyyy / filename)
' EJEMPLO: cscript FormatearFechas.vbs iso
' ==============================================================================

Option Explicit

Main()

Sub Main()
    Dim formatType
    If WScript.Arguments.Count > 0 Then
        formatType = LCase(WScript.Arguments(0))
    Else
        formatType = "iso"
    End If

    Select Case formatType
        Case "iso"
            ' YYYY-MM-DD
            WScript.StdOut.Write Year(Now) & "-" & Pad(Month(Now)) & "-" & Pad(Day(Now))
        Case "dd-mm-yyyy"
            WScript.StdOut.Write Pad(Day(Now)) & "-" & Pad(Month(Now)) & "-" & Year(Now)
        Case "filename"
            ' YYYYMMDD_HHMMSS
            WScript.StdOut.Write Year(Now) & Pad(Month(Now)) & Pad(Day(Now)) & "_" & Pad(Hour(Now)) & Pad(Minute(Now)) & Pad(Second(Now))
        Case Else
            WScript.StdOut.Write "FORMATO_NO_RECONOCIDO"
    End Select
End Sub

Function Pad(val)
    If len(val) = 1 Then
        Pad = "0" & val
    Else
        Pad = val
    End If
End Function
