Dim swApp As SldWorks.SldWorks

Function CreateBodiesFromSketches(vSketches As Variant, vDepths As Variant) As Variant
    
    Dim swBodies() As SldWorks.Body2
    
    If Not IsEmpty(vSketches) Then
            
        For i = 0 To UBound(vSketches)
            
            Dim swSketchFeat As SldWorks.Feature
            Set swSketchFeat = vSketches(i)
            
            Dim depth As Double
            depth = vDepths(i)
            
            Dim swSketch As SldWorks.sketch
            Set swSketch = swSketchFeat.GetSpecificFeature2
            
            Dim vSkRegs As Variant
            vSkRegs = swSketch.GetSketchRegions
            
            If Not IsEmpty(vSkRegs) Then
                
                Dim j As Integer
                
                For j = 0 To UBound(vSkRegs)
                
                    Dim swSkReg As SldWorks.SketchRegion
                    Set swSkReg = vSkRegs(j)
                    Dim swBody As SldWorks.Body2
                    Set swBody = Geometry.ExtrudeRegion(swSkReg, depth)
                    
                    If (Not swBodies) = -1 Then
                        ReDim swBodies(0)
                    Else
                        ReDim Preserve swBodies(UBound(swBodies) + 1)
                    End If
                    
                    Set swBodies(UBound(swBodies)) = swBody
                    
                Next
                
            End If
            
        Next
            
    End If
    
    If (Not swBodies) <> -1 Then
        CreateBodiesFromSketches = swBodies
    Else
        CreateBodiesFromSketches = Empty
    End If
    
End Function

Private Function ExtrudeRegion(swSkRegion As SldWorks.SketchRegion, depth As Double) As SldWorks.Body2
    
    Dim swBody As SldWorks.Body2
    Set swBody = CreateBodyFromRegion(swSkRegion)
    
    Set swApp = Application.SldWorks
    
    Dim swModeler As SldWorks.Modeler
    Set swModeler = swApp.GetModeler
    
    Dim swFace As SldWorks.Face2
    Set swFace = swBody.GetFaces()(0)
    Dim swDir As SldWorks.MathVector
    
    Dim swMathUtils As SldWorks.MathUtility
    Set swMathUtils = swApp.GetMathUtility
    
    Set swDir = swMathUtils.CreateVector(swFace.Normal)
    
    Set ExtrudeRegion = swModeler.CreateExtrudedBody(swBody, swDir, depth)
    
End Function

Private Function CreateBodyFromRegion(swSkRegion As SldWorks.SketchRegion) As SldWorks.Body2
            
    Set swApp = Application.SldWorks
    
    Dim vCurves As Variant
    vCurves = GetCurvesFromRegion(swSkRegion)
    
    Dim swSketch As SldWorks.sketch
    Set swSketch = swSkRegion.sketch
    
    Dim swSurf As SldWorks.Surface
    Set swSurf = CreatePlanarSurfaceFromSketch(swSketch)
    
    Dim swBody As SldWorks.Body2
    Set swBody = swSurf.CreateTrimmedSheet5(vCurves, True, 0.00001)
    
    Set CreateBodyFromRegion = swBody
    
End Function

Private Function GetCurvesFromRegion(skReg As SldWorks.SketchRegion) As Variant
    
    Dim swCurves() As SldWorks.Curve
    
    Dim swLoop As SldWorks.Loop2
    Set swLoop = skReg.GetFirstLoop()
    
    While Not swLoop Is Nothing
                
        Dim vLoopEdges As Variant
        vLoopEdges = swLoop.GetEdges
        
        If (Not swCurves) = -1 Then
            ReDim swCurves(UBound(vLoopEdges))
        Else
            If UBound(swCurves) = -1 Then
                ReDim swCurves(UBound(vLoopEdges))
            Else
                ReDim Preserve swCurves(UBound(swCurves) + UBound(vLoopEdges) + 2)
            End If
        End If
        
        Dim i As Integer
        
        For i = UBound(vLoopEdges) To 0 Step -1
            
            Dim swLoopEdge As SldWorks.Edge
            Set swLoopEdge = vLoopEdges(i)
            Dim swCurve As SldWorks.Curve
            Set swCurve = swLoopEdge.GetCurve().Copy
            Set swCurves(UBound(swCurves) - UBound(vLoopEdges) + i) = swCurve
            
        Next
        
        Set swLoop = swLoop.GetNext
        
    Wend
    
    GetCurvesFromRegion = swCurves
    
End Function

Private Function CreatePlanarSurfaceFromSketch(sketch As SldWorks.sketch) As SldWorks.Surface
    
    Dim swMathUtils As SldWorks.MathUtility
    Set swMathUtils = swApp.GetMathUtility
    
    Dim dPt(2) As Double
    Dim dVec(2) As Double
    
    Dim swRootPt As SldWorks.MathPoint
    dPt(0) = 0: dPt(1) = 0: dPt(2) = 0
    Set swRootPt = swMathUtils.CreatePoint(dPt)
    
    Dim swNormVec As SldWorks.MathVector
    dVec(0) = 0: dVec(1) = 0: dVec(2) = 1
    Set swNormVec = swMathUtils.CreateVector(dVec)
    
    Dim swRefVec As SldWorks.MathVector
    dVec(0) = 1: dVec(1) = 0: dVec(2) = 0
    Set swRefVec = swMathUtils.CreateVector(dVec)
    
    Dim swSkTransform As SldWorks.MathTransform
    Set swSkTransform = sketch.ModelToSketchTransform.Inverse
    
    Set swRootPt = swRootPt.MultiplyTransform(swSkTransform)
    Set swNormVec = swNormVec.MultiplyTransform(swSkTransform)
    Set swRefVec = swRefVec.MultiplyTransform(swSkTransform)
    
    Dim swModeler As SldWorks.Modeler
    Set swModeler = swApp.GetModeler
    
    Dim swSurf As SldWorks.Surface
    
    Set swSurf = swModeler.CreatePlanarSurface2(swRootPt.ArrayData, swNormVec.ArrayData, swRefVec.ArrayData)
    
    Set CreatePlanarSurfaceFromSketch = swSurf
    
End Function