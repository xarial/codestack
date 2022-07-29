Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swDraw As SldWorks.DrawingDoc
    
    Set swDraw = swApp.ActiveDoc
        
    Dim filePath As String
    filePath = swApp.GetOpenFileName("Select model to insert into a predefined views", "", _
        "SOLIDWORKS Model Files (*.sldprt; *.sldasm)|*.sldprt;*.sldasm|All Files (*.*)|*.*|", 0, "", "")
    
    If filePath <> "" Then
    
        If False = swDraw.InsertModelInPredefinedView(filePath) Then
            Err.Raise vbError, "", "Failed to insert model into predefined views"
        End If
    
    End If
    
End Sub