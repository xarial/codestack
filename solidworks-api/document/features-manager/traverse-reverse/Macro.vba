Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swModel As SldWorks.ModelDoc2
    
    Set swModel = swApp.ActiveDoc
    
    Dim i As Integer
    
    i = 0
    
    Dim swFeat As SldWorks.Feature
    
    Do
        
        Set swFeat = swModel.FeatureByPositionReverse(i)
        i = i + 1
        
        If Not swFeat Is Nothing Then
            Debug.Print swFeat.Name
        End If
        
    Loop While Not swFeat Is Nothing
    
End Sub