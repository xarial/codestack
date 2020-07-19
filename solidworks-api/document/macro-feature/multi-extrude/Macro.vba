Dim swController As Controller

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swModel As SldWorks.ModelDoc2
    
    Set swModel = swApp.ActiveDoc
    
    If Not swModel Is Nothing Then
        
        Set swController = New Controller
        swController.InsertExtrude
        
    Else
        MsgBox "Please open model"
    End If
    
End Sub