Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    Dim swPart As SldWorks.PartDoc
    
    Set swPart = swApp.ActiveDoc
    
    If Not swPart Is Nothing Then
            
        Dim vBBox As Variant
    
        vBBox = GetPreciseBoundingBox(swPart)
     
        DrawBox swPart, CDbl(vBBox(0)), CDbl(vBBox(1)), CDbl(vBBox(2)), CDbl(vBBox(3)), CDbl(vBBox(4)), CDbl(vBBox(5))
        
        Debug.Print "Width: " & maxX - minX
        Debug.Print "Length: " & maxZ - minZ
        Debug.Print "Height: " & maxY - minY
        
    Else
        
        MsgBox "Please open part"
        
    End If
    
End Sub

Function GetPreciseBoundingBox(part As SldWorks.PartDoc) As Variant
    
    Dim dBox(5) As Double
    
    Dim vBodies As Variant
    vBodies = part.GetBodies2(swBodyType_e.swSolidBody, True)
        
    Dim minX As Double
    Dim minY As Double
    Dim minZ As Double
    Dim maxX As Double
    Dim maxY As Double
    Dim maxZ As Double
        
    If Not IsEmpty(vBodies) Then
    
        Dim i As Integer
        
        For i = 0 To UBound(vBodies)
        
            Dim swBody As SldWorks.Body2
    
            Set swBody = vBodies(i)
            
            Dim x As Double
            Dim y As Double
            Dim z As Double
            
            swBody.GetExtremePoint 1, 0, 0, x, y, z
            
            If i = 0 Or x > maxX Then
                maxX = x
            End If
            
            swBody.GetExtremePoint -1, 0, 0, x, y, z
            
            If i = 0 Or x < minX Then
                minX = x
            End If
            
            swBody.GetExtremePoint 0, 1, 0, x, y, z
            
            If i = 0 Or y > maxY Then
                maxY = y
            End If
            
            swBody.GetExtremePoint 0, -1, 0, x, y, z
            
            If i = 0 Or y < minY Then
                minY = y
            End If
            
            swBody.GetExtremePoint 0, 0, 1, x, y, z
            
            If i = 0 Or z > maxZ Then
                maxZ = z
            End If
            
            swBody.GetExtremePoint 0, 0, -1, x, y, z
            
            If i = 0 Or z < minZ Then
                minZ = z
            End If
            
        Next
    
    End If
    
    dBox(0) = minX: dBox(1) = minY: dBox(2) = minZ
    dBox(3) = maxX: dBox(4) = maxY: dBox(5) = maxZ
    
    GetPreciseBoundingBox = dBox
    
End Function

Sub DrawBox(model As SldWorks.ModelDoc2, minX As Double, minY As Double, minZ As Double, maxX As Double, maxY As Double, maxZ As Double)

    model.ClearSelection2 True
            
    model.SketchManager.Insert3DSketch True
    model.SketchManager.AddToDB = True
    
    model.SketchManager.CreateLine maxX, minY, minZ, maxX, minY, maxZ
    model.SketchManager.CreateLine maxX, minY, maxZ, minX, minY, maxZ
    model.SketchManager.CreateLine minX, minY, maxZ, minX, minY, minZ
    model.SketchManager.CreateLine minX, minY, minZ, maxX, minY, minZ

    model.SketchManager.CreateLine maxX, maxY, minZ, maxX, maxY, maxZ
    model.SketchManager.CreateLine maxX, maxY, maxZ, minX, maxY, maxZ
    model.SketchManager.CreateLine minX, maxY, maxZ, minX, maxY, minZ
    model.SketchManager.CreateLine minX, maxY, minZ, maxX, maxY, minZ
    
    model.SketchManager.CreateLine minX, minY, minZ, minX, maxY, minZ
    model.SketchManager.CreateLine minX, minY, maxZ, minX, maxY, maxZ
    
    model.SketchManager.CreateLine maxX, minY, minZ, maxX, maxY, minZ
    model.SketchManager.CreateLine maxX, minY, maxZ, maxX, maxY, maxZ
    
    model.SketchManager.AddToDB = False
    model.SketchManager.Insert3DSketch True
    
End Sub