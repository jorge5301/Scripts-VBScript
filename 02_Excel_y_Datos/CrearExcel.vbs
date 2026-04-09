' ==============================================================================
' NOMBRE: CrearExcel.vbs
' DESCRIPCIÓN: Crea un nuevo libro de Excel (.xlsx) en la ruta especificada.
' ARGUMENTOS: 1. Ruta completa del archivo a crear.
' EJEMPLO: cscript CrearExcel.vbs "C:\Temp\Reporte.xlsx"
' ==============================================================================

Option Explicit

Main()

Sub Main()
    If WScript.Arguments.Count < 1 Then
        WScript.Echo "ERROR: Se requiere la ruta completa como argumento."
        WScript.Quit(1)
    End If

    Dim rutaArchivo, fso
    rutaArchivo = WScript.Arguments(0)
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Validar si el directorio existe, si no, intentar crearlo
    Dim carpeta
    carpeta = fso.GetParentFolderName(rutaArchivo)
    If Not fso.FolderExists(carpeta) Then
        On Error Resume Next
        CrearCarpetaRecursiva carpeta
        If Err.Number <> 0 Then
            WScript.Echo "ERROR: No se pudo crear el directorio: " & carpeta
            WScript.Quit(1)
        End If
        On Error Goto 0
    End If

    ' Crear el Excel
    Dim xlApp, xlBook
    On Error Resume Next
    Set xlApp = CreateObject("Excel.Application")
    If Err.Number <> 0 Then
        WScript.Echo "ERROR: No se pudo iniciar Excel. Verifique la instalación."
        WScript.Quit(1)
    End If

    xlApp.DisplayAlerts = False
    Set xlBook = xlApp.Workbooks.Add
    
    ' 51 = xlOpenXMLWorkbook (xlsx)
    xlBook.SaveAs rutaArchivo, 51
    
    If Err.Number <> 0 Then
        WScript.Echo "ERROR: No se pudo guardar el archivo en: " & rutaArchivo
        xlApp.Quit
        Set xlApp = Nothing
        WScript.Quit(1)
    End If

    xlBook.Close False
    xlApp.Quit
    
    Set xlBook = Nothing
    Set xlApp = Nothing
    Set fso = Nothing

    WScript.Echo "ÉXITO: Archivo creado en " & rutaArchivo
End Sub

Sub CrearCarpetaRecursiva(ByVal ruta)
    Dim fso, padre
    Set fso = CreateObject("Scripting.FileSystemObject")
    padre = fso.GetParentFolderName(ruta)
    If Not fso.FolderExists(padre) Then
        CrearCarpetaRecursiva padre
    End If
    If Not fso.FolderExists(ruta) Then
        fso.CreateFolder(ruta)
    End If
End Sub
