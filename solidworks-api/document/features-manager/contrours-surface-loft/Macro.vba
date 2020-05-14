Dim swApp As SldWorks.SldWorks

Sub main()

    Dim swModel As SldWorks.ModelDoc2
    Dim swSelMgr As SldWorks.SelectionMgr
    
    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc

    Set swSelMgr = swModel.SelectionManager
    
    Dim swContours() As SldWorks.SketchContour
    ReDim swContours(swSelMgr.GetSelectedObjectCount2(-1) - 1)
    
    Dim i As Integer
    
    For i = 1 To swSelMgr.GetSelectedObjectCount2(-1)
        Dim swSkSeg As SldWorks.SketchSegment
        Set swSkSeg = swSelMgr.GetSelectedObject6(i, -1)
        Set swContours(i - 1) = GetSketchContour(swSkSeg)
    Next
    
    swModel.ClearSelection2 True
    
    Dim swSelData As SldWorks.SelectData
        
    Set swSelData = swSelMgr.CreateSelectData
    
    swSelData.Mark = 1
        
    For i = 0 To UBound(swContours)
        Dim swSkContour As SldWorks.SketchContour
        Set swSkContour = swContours(i)
        swSkContour.Select2 True, swSelData
    Next
    
    swModel.InsertLoftRefSurface2 False, True, False, 1, 0, 0

End Sub

Function GetSketchContour(sketchSeg As SldWorks.SketchSegment) As SldWorks.SketchContour
    
    Dim swSketch As SldWorks.Sketch
    Set swSketch = sketchSeg.GetSketch
    
    Dim vSketchContours As Variant
    
    vSketchContours = swSketch.GetSketchContours
    
    If Not IsEmpty(vSketchContours) Then
        
        Dim i As Integer
        
        For i = 0 To UBound(vSketchContours)
            
            Dim swSkContour As SldWorks.SketchContour
            Set swSkContour = vSketchContours(i)
            
            Dim vSegs As Variant
            vSegs = swSkContour.GetSketchSegments()
            
            If Not IsEmpty(vSegs) Then
                
                Dim j As Integer
                
                Dim swCurSkSeg As SldWorks.SketchSegment
                Set swCurSkSeg = vSegs(j)
                
                If swApp.IsSame(sketchSeg, swCurSkSeg) = swObjectEquality.swObjectSame Then
                    Set GetSketchContour = swSkContour
                    Exit Function
                End If
                
            End If
            
        Next
        
    End If
    
End Function