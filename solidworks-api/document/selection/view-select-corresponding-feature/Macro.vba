Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swModel As SldWorks.ModelDoc2
    
    Set swModel = swApp.ActiveDoc
    
    Dim swFeat As SldWorks.Feature
    
    Dim swSelMgr As SldWorks.SelectionMgr
    
    Set swSelMgr = swModel.SelectionManager
    
    Set swFeat = swSelMgr.GetSelectedObject6(1, -1)
    
    'activate drawing
    Stop
    
    Dim swDraw As SldWorks.DrawingDoc
    
    Set swDraw = swApp.ActiveDoc
        
    Set swSelMgr = swDraw.SelectionManager
    
    Dim vViews As Variant
    
    vViews = swDraw.GetViews()(0)
    
    Dim i As Integer
    
    Dim swSelData As SldWorks.SelectData
    Set swSelData = swSelMgr.CreateSelectData
    
    swDraw.ClearSelection2 True
    
    For i = 0 To UBound(vViews)
        
        Dim swView As SldWorks.View
        
        Set swView = vViews(i)
        
        If swView.ReferencedDocument Is swModel Then
                    
            Dim swViewFeat As SldWorks.Entity
            Set swViewFeat = swFeat
            
            Set swViewFeat = swView.GetCorresponding(swFeat)
            
            swSelData.View = swView
            
            If Not swViewFeat Is Nothing Then
                Debug.Print swViewFeat.Select4(True, swSelData)
            Else
                Debug.Print "Failed to get corresponding feature"
            End If
            
        End If
        
    Next
    
End Sub