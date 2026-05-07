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
    
        Dim swLoop As SldWorks.Loop2
        Set swLoop = swFace.GetFirstLoop
        
        While Not swLoop Is Nothing
            
            swModel.ClearSelection2 True
            
            Dim isOuter As Boolean
            
            isOuter = False <> swLoop.isOuter()
            
            If isOuter Then
                Debug.Print "Traversing outer loop (CCW)"
            Else
                Debug.Print "Traversing inner loop (CW)"
            End If
            
            Dim vEdges As Variant
            vEdges = swLoop.GetEdges()
            
            Dim i As Integer
            
            For i = 0 To UBound(vEdges)
            
                Dim swEdge As SldWorks.Edge
                Set swEdge = vEdges(i)
                
                Dim swPrevPt As Vertex
                Dim swNextPt As Vertex
                
                If False <> swEdge.EdgeInFaceSense(swFace) Then
                    Set swPrevPt = swEdge.GetStartVertex()
                    Set swNextPt = swEdge.GetEndVertex()
                Else
                    Set swNextPt = swEdge.GetStartVertex()
                    Set swPrevPt = swEdge.GetEndVertex()
                End If
                
                If i = 0 Then
                    swPrevPt.Select4 True, Nothing
                    Stop
                End If
                
                If i <> UBound(vEdges) Then
                    swNextPt.Select4 True, Nothing
                    Stop
                End If
                
            Next
        
            Set swLoop = swLoop.GetNext
        Wend
    
    Else
        Err.Raise vbError, "", "Select face"
    End If
    
End Sub