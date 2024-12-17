Public Sub EnviarCorreoOutlook()
' *******************************************************************************************
'              Esta función envia un correo y adjunta una hoja exportada a PDF
' *******************************************************************************************    
    
    ' Declaro variables
    Dim OutlookApp As Object
    Dim Correo As Object
    Dim RutaArhivo As String
    
    ' Inicializo las variables
    RutaArhivo = "C:\..."
    
    ' Creo objeto Outlook
    Set OutlookApp = CreateObject("Outlook.Application")
    Set Correo = OutlookApp.CreateItem(0)
    
    ' Datos en el correo
    With Correo
        .To = "..."
        .Subject = "..." & Format(Date, "dd-mm-yyyy")
        .Body = "..."
        .Attachments.Add RutaArhivo
        .Display ' Muestra la ventana del correo para revisar
        '.Send ' Envia el correo automáticamente sin revision
    End With
    
    ' Limpiar
    Set Correo = Nothing
    Set OutlookApp = Nothing

End Sub
