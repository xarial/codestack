Dim swApp As SldWorks.SldWorks
Const HIT_RADIUS As Double = 0.00000001

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swModel As SldWorks.ModelDoc2
    Dim swPart As SldWorks.PartDoc
    
    Set swModel = swApp.ActiveDoc
    Set swPart = swModel
    
    Dim swSelMgr As SldWorks.SelectionMgr
    Set swSelMgr = swModel.SelectionManager
    
    Dim swSketch As SldWorks.sketch
    
    If swSelMgr.GetSelectedObjectType3(1, -1) = swSelectType_e.swSelSKETCHES Then
        Dim swFeat As SldWorks.Feature
        Set swFeat = swSelMgr.GetSelectedObject6(1, -1)
        Set swSketch = swFeat.GetSpecificFeature2
    Else
        Err.Raise vbError, "", "Sketch with sketch point rays is not selected"
    End If
    
    Dim vRayStartPts As Variant
    Dim vRayVecs As Variant
    
    GetRaysFromSketchPoints swSketch, vRayStartPts, vRayVecs
    
    Dim vBodies As Variant
    vBodies = swPart.GetBodies2(swBodyType_e.swSolidBody, True)
    
    Dim interCount As Integer
    interCount = swModel.Extension.RayIntersections(vBodies, vRayStartPts, vRayVecs, swRayPtsOpts_e.swRayPtsOptsENTRY_EXIT + swRayPtsOpts_e.swRayPtsOptsTOPOLS, HIT_RADIUS, 0, True)
    
    If interCount > 0 Then
        
        Dim vInterPoints As Variant
        vInterPoints = swModel.GetRayIntersectionsPoints()
        
        Dim vInterTopol As Variant
        vInterTopol = swModel.GetRayIntersectionsTopology
        
        Dim i As Integer
        
        For i = 0 To interCount - 1
            
            Dim bodyIndex As Integer
            Dim rayIndex As Integer
            Dim interType As Integer
            Dim dHitPt(2) As Double
            
            bodyIndex = CInt(vInterPoints(i * 9))
            rayIndex = CInt(vInterPoints(i * 9 + 1))
            interType = CInt(vInterPoints(i * 9 + 2))
            
            dHitPt(0) = CDbl(vInterPoints(i * 9 + 3))
            dHitPt(1) = CDbl(vInterPoints(i * 9 + 4))
            dHitPt(2) = CDbl(vInterPoints(i * 9 + 5))
            
            Dim swEnt As SldWorks.Entity
            Set swEnt = vInterTopol(i)
            
            Debug.Print "Intersecting body: " & vBodies(bodyIndex).Name
            Debug.Print "Intersecting ray: [" & vRayStartPts(rayIndex * 3) & ";" & vRayStartPts(rayIndex * 3 + 1) & ";" & vRayStartPts(rayIndex * 3 + 2) & "] - [" & vRayVecs(rayIndex * 3) & ";" & vRayVecs(rayIndex * 3 + 1) & ";" & vRayVecs(rayIndex * 3 + 2) & "]"
            Debug.Print "Intersection type: " & interType
            
            Dim swSelData As SldWorks.SelectData
            Set swSelData = swSelMgr.CreateSelectData
            
            swSelData.X = dHitPt(0)
            swSelData.Y = dHitPt(1)
            swSelData.Z = dHitPt(2)
            
            swEnt.Select4 False, swSelData
            
            Stop
            
        Next
        
    Else
        Err.Raise vbError, "", "No intersections found"
    End If
    
End Sub

Sub GetRaysFromSketchPoints(sketch As SldWorks.sketch, rayStartPts As Variant, rayVecs As Variant)
    
    If False = sketch.Is3D() Then
        
        Dim dRayStartPts() As Double
        Dim dRayVecs() As Double
        
        Dim vSkPoints As Variant
        vSkPoints = sketch.GetSketchPoints2
        
        If Not IsEmpty(vSkPoints) Then
            
            Dim swTransform As SldWorks.MathTransform
            Set swTransform = sketch.ModelToSketchTransform.Inverse
            
            Dim swMathUtils As SldWorks.MathUtility
            Set swMathUtils = swApp.GetMathUtility
            
            Dim dVec(2) As Double
            dVec(0) = 0: dVec(1) = 0: dVec(2) = 1
            
            Dim swMathVec As SldWorks.MathVector
            Set swMathVec = swMathUtils.CreateVector(dVec)
            Set swMathVec = swMathVec.MultiplyTransform(swTransform)
            
            ReDim dRayStartPts((UBound(vSkPoints) + 1) * 3 - 1)
            ReDim dRayVecs((UBound(vSkPoints) + 1) * 3 - 1)
            
            Dim i As Integer
            
            For i = 0 To UBound(vSkPoints)
                
                Dim swMathPt As SldWorks.MathPoint
                Dim dPt(2) As Double
                
                Dim swSkPt As SldWorks.SketchPoint
                Set swSkPt = vSkPoints(i)
                dPt(0) = swSkPt.X: dPt(1) = swSkPt.Y: dPt(2) = 0
                
                Set swMathPt = swMathUtils.CreatePoint(dPt)
                Set swMathPt = swMathPt.MultiplyTransform(swTransform)
                
                Dim vData As Variant
                vData = swMathPt.ArrayData
                
                dRayStartPts(i * 3) = vData(0)
                dRayStartPts(i * 3 + 1) = vData(1)
                dRayStartPts(i * 3 + 2) = vData(2)
                
                vData = swMathVec.ArrayData
                
                dRayVecs(i * 3) = vData(0)
                dRayVecs(i * 3 + 1) = vData(1)
                dRayVecs(i * 3 + 2) = vData(2)
                
            Next
            
            rayStartPts = dRayStartPts
            rayVecs = dRayVecs
            
        Else
            Err.Raise vbError, "", "No sketch points in the specified sketch"
        End If
        
    Else
        Err.Raise vbError, "", "Only 2D sketch can be used for rays"
    End If
    
End Sub