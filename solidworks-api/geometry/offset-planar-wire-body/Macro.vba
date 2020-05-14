Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swModel As SldWorks.ModelDoc2
    
    Set swModel = swApp.ActiveDoc
    
    If Not swModel Is Nothing Then
        
        Dim swSelMgr As SldWorks.SelectionMgr
        Set swSelMgr = swModel.SelectionManager
        
        Dim swEdge As SldWorks.Edge
        Set swEdge = swSelMgr.GetSelectedObject6(1, -1)
        
        If Not swEdge Is Nothing Then
        
            Dim swBody As SldWorks.Body2
            Set swBody = swEdge.GetBody()
            
            If swBody.GetType() = swBodyType_e.swWireBody Then
                
                Dim swOffsetBody As SldWorks.Body2
                Dim swNormVec As SldWorks.MathVector
                
                Dim swMathUtils As SldWorks.MathUtility
                Set swMathUtils = swApp.GetMathUtility
                
                Dim dVec(2) As Double
                dVec(0) = 0: dVec(1) = 0: dVec(2) = 1
                
                Set swNormVec = swMathUtils.CreateVector(dVec)
                
                Set swOffsetBody = swBody.OffsetPlanarWireBody(0.01, swNormVec, swOffsetPlanarWireBodyOptions_e.swOffsetPlanarWireBodyOptions_GapFillExtend)
                
                If swOffsetBody Is Nothing Then
                    Err.Raise vbError, "", "Failed to create offset body. Make sure that selected edge is on a plane with the normal specified in dVec variable"
                End If
                
                swOffsetBody.Display3 swModel, RGB(255, 255, 0), swTempBodySelectOptions_e.swTempBodySelectOptionNone
                
                Stop
                
                Set swOffsetBody = Nothing
                
            Else
                Err.Raise vbError, "", "Selected edge is not a wire body"
            End If
        
        Else
            Err.Raise "Edge is not selected"
        End If
        
    Else
        Err.Raise "Document is not open"
    End If
    
End Sub