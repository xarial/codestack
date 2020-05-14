Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swModel As SldWorks.ModelDoc2
    
    Set swModel = swApp.ActiveDoc
    
    If Not swModel Is Nothing Then
        
        Dim swSkArc As SldWorks.SketchArc
        Set swSkArc = swModel.SelectionManager.GetSelectedObject6(1, -1)
        
        If Not swSkArc Is Nothing Then
            Dim swEndPts(1) As SldWorks.SketchPoint
            Set swEndPts(0) = swSkArc.GetStartPoint2()
            Set swEndPts(1) = swSkArc.GetEndPoint2()
            swModel.SketchManager.ActiveSketch.RelationManager.AddRelation swEndPts, swConstraintType_e.swConstraintType_MERGEPOINTS
        Else
            MsgBox "Please select sketch arc"
        End If
        
    Else
        MsgBox "Please open the model"
    End If
    
End Sub