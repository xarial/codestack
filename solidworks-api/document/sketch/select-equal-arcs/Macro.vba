Const EPS As Double = 0.0000000001

Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    On Error GoTo catch
    
try:
    Dim swModel As SldWorks.ModelDoc2
    
    Set swModel = swApp.ActiveDoc
    
    If Not swModel Is Nothing Then
        
        Dim swSkSrcArc As SldWorks.SketchArc
        Set swSkSrcArc = swModel.SelectionManager.GetSelectedObject6(1, -1)
        
        If Not swSkSrcArc Is Nothing Then
            
            Dim radius As Double
            radius = swSkSrcArc.GetRadius()
            
            Dim swSketch As SldWorks.Sketch
            Set swSketch = swSkSrcArc.GetSketch
            
            Dim vSegs As Variant
            vSegs = swSketch.GetSketchSegments()
            
            Dim i As Integer
            
            For i = 0 To UBound(vSegs)
                
                Dim swSkSeg As SldWorks.SketchSegment
                Set swSkSeg = vSegs(i)
                
                If swSkSeg.GetType() = swSketchSegments_e.swSketchARC Then
                
                    If Not swSkSrcArc Is swSkSeg Then
                    
                        Dim swSkArc As SldWorks.SketchArc
                        Set swSkArc = swSkSeg
                        
                        If Abs(swSkArc.GetRadius() - radius) < EPS Then
                            Dim swSkArcSeg As SldWorks.SketchSegment
                            Set swSkArcSeg = swSkArc
                            swSkArcSeg.Select4 True, Nothing
                        End If
                        
                    End If
                End If
                
            Next
            
        Else
            Err.Raise vbError, "", "Please select sketch arc"
        End If
        
    Else
        Err.Raise vbError, "", "Open model"
    End If
    
    GoTo finally
catch:
    swApp.SendMsgToUser2 Err.Description, swMessageBoxIcon_e.swMbStop, swMessageBoxBtn_e.swMbOk
finally:
    
End Sub