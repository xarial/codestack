Dim swApp As SldWorks.SldWorks
Dim swModel As SldWorks.ModelDoc2

Sub main() 'this method is the only one without parameters

    ConnectToSw Empty
    CountFeatures Empty
    
End Sub

Sub ConnectToSw(dummy)
    
    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    
    If swModel Is Nothing Then
        MsgBox "Please open the model"
        End
    End If
    
End Sub

Sub CountFeatures(dummy) 'this method could be selected as an entry point
    
    swApp.SendMsgToUser "There are " & swModel.GetFeatureCount() & " features in the active model"
    
End Sub