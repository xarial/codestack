Dim swApp As SldWorks.SldWorks
Dim swModel As SldWorks.ModelDoc2

Sub main() 'this method must be an entry point

    ConnectToSw
    CountFeatures
    
End Sub

Sub ConnectToSw() 'this method could be selected as an entry point
    
    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    
    If swModel Is Nothing Then
        MsgBox "Please open the model"
        End
    End If
    
End Sub

Sub CountFeatures() 'this method could be selected as an entry point
    
    swApp.SendMsgToUser "There are " & swModel.GetFeatureCount() & " features in the active model"
    
End Sub