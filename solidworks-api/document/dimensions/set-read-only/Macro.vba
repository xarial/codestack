Const READ_ONLY As Boolean = True

Dim swApp As SldWorks.SldWorks
Dim swModel As SldWorks.ModelDoc2

Sub main()

    Set swApp = Application.SldWorks
    
    Set swModel = swApp.ActiveDoc
    
    Dim swSelMgr As SldWorks.SelectionMgr
    
    Set swSelMgr = swModel.SelectionManager
    
    Dim swFeat As SldWorks.Feature
    
    Set swFeat = swSelMgr.GetSelectedObject6(1, -1)
    
    If Not swFeat Is Nothing Then
        
        Dim swDispDim As SldWorks.DisplayDimension
        
        Set swDispDim = swFeat.GetFirstDisplayDimension
        
        While Not swDispDim Is Nothing
            
            Dim swDim As SldWorks.Dimension
            
            Set swDim = swDispDim.GetDimension2(0)
            swDim.ReadOnly = READ_ONLY
            
            Set swDispDim = swFeat.GetNextDisplayDimension(swDispDim)
            
        Wend
        
    Else
        Err.Raise vbError, "", "Select feature"
    End If
    
End Sub