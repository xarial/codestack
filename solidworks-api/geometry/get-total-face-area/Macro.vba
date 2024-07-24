Const VISIBLE_ONLY As Boolean = True

Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swPart As SldWorks.PartDoc
    
    Set swPart = swApp.ActiveDoc
    
    If Not swPart Is Nothing Then
    
        Dim vBodies As Variant
        
        vBodies = swPart.GetBodies2(swBodyType_e.swAllBodies, VISIBLE_ONLY)
        
        Dim totalArea As Double
        
        If Not IsEmpty(vBodies) Then
            
            Dim i As Integer
            
            For i = 0 To UBound(vBodies)
                
                Dim swBody As SldWorks.Body2
                
                Set swBody = vBodies(i)
                Dim vFaces As Variant
                
                vFaces = swBody.GetFaces
                
                If Not IsEmpty(vFaces) Then
                    Dim j As Integer
                    
                    For j = 0 To UBound(vFaces)
                        Dim swFace As SldWorks.Face2
                        Set swFace = vFaces(j)
                        totalArea = totalArea + swFace.GetArea
                    Next
                    
                End If
                
            Next
            
            Dim swUserUnit As SldWorks.UserUnit
            Set swUserUnit = swPart.GetUserUnit(swUserUnitsType_e.swLengthUnit)
            
            Dim convFactor As Double
            convFactor = swUserUnit.GetConversionFactor() ^ 2
            
            MsgBox "Total Area: " & totalArea * convFactor & " " & swUserUnit.GetUnitsString(False) & "^2"
            
        Else
            Err.Raise vbError, "", "No bodies"
        End If
    
    Else
        Err.Raise vbError, "", "Open part file"
    End If
    
End Sub