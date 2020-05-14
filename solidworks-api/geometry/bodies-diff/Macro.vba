Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swPart As SldWorks.PartDoc
    
    Set swPart = swApp.ActiveDoc
    
    If Not swPart Is Nothing Then
        
        Dim otherFilePath As String
        otherFilePath = InputBox("Please specify the part path to compare to")
        
        If otherFilePath <> "" Then
            
            Dim swOtherPart As SldWorks.PartDoc
            Set swOtherPart = swApp.OpenDoc6(otherFilePath, swDocumentTypes_e.swDocPART, swOpenDocOptions_e.swOpenDocOptions_Silent, "", 0, 0)
            
            If Not swOtherPart Is Nothing Then
                Dim swXform As SldWorks.MathTransform
                Set swXform = GetClosestTransform(swPart, swOtherPart)
                PreviewPart swOtherPart, swXform, swPart
            Else
                MsgBox "Failed to open the part to compare to"
            End If
            
        End If
        
    Else
        MsgBox "Please open part"
    End If
        
End Sub

Sub PreviewPart(part As SldWorks.PartDoc, transform As SldWorks.MathTransform, context As PartDoc)
    
    Dim vBodies As Variant
    
    vBodies = part.GetBodies2(swBodyType_e.swSolidBody, True)
    
    Dim i As Integer
    
    For i = 0 To UBound(vBodies)
        
        Dim swBody As SldWorks.Body2
        Set swBody = vBodies(i)
        Set swBody = swBody.Copy
        
        If Not transform Is Nothing Then
            Debug.Print swBody.ApplyTransform(transform)
        End If
        
        Set vBodies(i) = swBody
        swBody.Display3 context, RGB(255, 255, 0), swTempBodySelectOptions_e.swTempBodySelectOptionNone
        
    Next
    
    Stop 'continue the macro to hide preview
    
End Sub

Function GetClosestTransform(thisPart As SldWorks.PartDoc, otherPart As SldWorks.PartDoc) As SldWorks.MathTransform

    Dim vThisBodies As Variant
    Dim vOtherBodies As Variant
    
    vThisBodies = thisPart.GetBodies2(swBodyType_e.swSolidBody, True)
    vOtherBodies = otherPart.GetBodies2(swBodyType_e.swSolidBody, True)
    
    Dim transformsHits As Object
    Set transformsHits = CreateObject("Scripting.Dictionary") 'by some reasons sometimes the first null element is added on creation
        
    If Not IsEmpty(vThisBodies) And Not IsEmpty(vOtherBodies) Then
        
        Dim i As Integer
        Dim j As Integer
        
        For i = 0 To UBound(vOtherBodies)
            
            Dim swOtherBody As SldWorks.Body2
                        
            Set swOtherBody = vOtherBodies(i)
            
            For j = 0 To UBound(vThisBodies)
                
                Dim swThisBody As SldWorks.Body2
                Set swThisBody = vThisBodies(j)
                Dim swTransform As SldWorks.MathTransform
                
                If swThisBody.GetCoincidenceTransform2(swOtherBody, swTransform) Then
                    
                    If Not swTransform Is Nothing Then
                        
                        Dim contains As Boolean
                        contains = False
                        For Each key In transformsHits.Keys
                            If Not key Is Nothing Then
                                Dim tx As SldWorks.MathTransform
                                Set tx = key
                                If CompareTransforms(swTransform, tx) Then
                                    transformsHits(tx) = transformsHits(tx) + 1
                                    contains = True
                                    Exit For
                                End If
                            End If
                        Next
                        
                        If Not contains Then
                            transformsHits.Add swTransform, 1
                        End If
                        
                    End If
                    
                End If
                
            Next
            
        Next
        
    End If
    
    Dim curMaxHit As Integer
    curMaxHit = 0
    
    For Each key In transformsHits.Keys
        If Not key Is Nothing Then
            Dim curTx As SldWorks.MathTransform
            Set curTx = key
            If transformsHits(curTx) > curMaxHit Then
                curMaxHit = transformsHits(curTx)
                Set GetClosestTransform = curTx
            End If
        End If
    Next

End Function

Function CompareTransforms(firstTransform As SldWorks.MathTransform, secondTransform As SldWorks.MathTransform) As Boolean
    
    Dim vFirstArrayData As Variant
    vFirstArrayData = firstTransform.ArrayData
    
    Dim vSecondArrayData As Variant
    vSecondArrayData = secondTransform.ArrayData
    
    Dim i As Integer
    
    For i = 0 To UBound(vFirstArrayData)
        If Not CompareValues(CDbl(vFirstArrayData(i)), CDbl(vSecondArrayData(i))) Then
            CompareTransforms = False
            Exit Function
        End If
    Next
    
    CompareTransforms = True
    
End Function

Function CompareValues(firstValue As Double, secondValue As Double, Optional tol As Double = 0.00000001) As Boolean
        
    CompareValues = Abs(secondValue - firstValue) <= tol
    
End Function