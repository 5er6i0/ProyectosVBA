Public Sub CrearPDF()
' *******************************************************************************************
'                   Esta funcion exporta una hoja de excel a PDF
' *******************************************************************************************
' Nota: suguiero combinarlo con la funcion de area de impresion para mejor presentacion 

    ' Declaro Variables
    Dim wsNombreHoja As Worksheet
    Dim RutaTemp As String
    
    ' Inicializo variables
    Set wsNombreHoja = ThisWorkbook.Sheets("Nombre de la hoja")
    RutaTemp = Environ("TEMP")

    ' Generamos la hoja como PDF y la almacenamos como un archivo temporal
    Application.DisplayAlerts = False
    RutaPDF = RutaTemp & "\" & "Prueba" & Format(Now, "ddmmyy-hhmmss") & ".pdf"
    wsNombreHoja.ExportAsFixedFormat Type:=xlTypePDF, Filename:=RutaPDF, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
    
    ' Abrimos el PDF generado
    Shell "cmd /c start " & RutaPDF, vbNormalFocus

End Sub