Const MIN_ACCURACY As Long = 1
Const MAX_ACCURACY As Long = 10

Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swModel As SldWorks.ModelDoc2
    
    Set swModel = swApp.ActiveDoc
    
    Dim swSelMgr As SldWorks.SelectionMgr
    
    Set swSelMgr = swModel.SelectionManager
    
    Dim swFace As SldWorks.Face2
    
    Set swFace = swSelMgr.GetSelectedObject6(1, -1)
    
    If Not swFace Is Nothing Then
        
        Dim swEdge As SldWorks.Edge
        Set swEdge = swFace.GetEdges()(0)
        Dim swVertex As SldWorks.Vertex
        Set swVertex = swEdge.GetStartVertex
        
        swFace.SelectByMark False, 1
        swVertex.SelectByMark True, 16
        
        Dim swFlatSurfFeat As SldWorks.Feature
        Set swFlatSurfFeat = swModel.FeatureManager.InsertFlattenSurface2(CLng((MIN_ACCURACY + MAX_ACCURACY) / 2), False)
        
        If swFlatSurfFeat Is Nothing Then
            Err.Raise vbError, "", "Failed to create flatten surface"
        End If
        
    Else
        Err.Raise vbError, "", "Select face"
    End If
    
End Sub