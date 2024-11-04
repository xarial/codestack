Enum ExtendSurfaceEndCondition_e
    Distance = 0
    UpToVertex = 1
    UpToFace = 2
End Enum

Const EXTEND_DISTANCE As Double = 0.01

Dim swApp As SldWorks.SldWorks

Sub main()
    
    Set swApp = Application.SldWorks
    
    Dim swModel As SldWorks.ModelDoc2
    Dim swSelMgr As SldWorks.SelectionMgr
    
    Set swModel = swApp.ActiveDoc
    
    If Not swModel Is Nothing Then
        
        Set swSelMgr = swModel.SelectionManager
        
        Dim swFace As SldWorks.Face2
        Set swFace = swSelMgr.GetSelectedObject6(1, -1)
        
        If Not swFace Is Nothing Then
        
            Dim swBody As SldWorks.Body2
            
            Set swBody = swFace.CreateSheetBody
                
            Dim vEdges As Variant
            vEdges = swBody.GetEdges
            
            Dim swExtendedBody As SldWorks.Body2
            Set swExtendedBody = swBody.ExtendSurface(vEdges, True, ExtendSurfaceEndCondition_e.Distance, EXTEND_DISTANCE, Nothing, Nothing)
            
            If Not swExtendedBody Is Nothing Then
                swExtendedBody.Display2 swModel, RGB(255, 255, 0), swTempBodySelectOptions_e.swTempBodySelectOptionNone
                Stop
                Set swExtendedBody = Nothing
            Else
                Err.Raise vbError, "", "Failed to extend the selected face"
            End If
            
        Else
            Err.Raise vbError, "", "Select face to extend"
        End If
    
    Else
        Err.Raise vbError, "", "Open part document"
    End If
    
End Sub