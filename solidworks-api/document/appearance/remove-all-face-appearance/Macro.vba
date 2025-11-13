Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swModel As SldWorks.ModelDoc2
    
    Set swModel = swApp.ActiveDoc
    
    Dim vAppearances As Variant
    
    vAppearances = swModel.Extension.GetRenderMaterials2(swDisplayStateOpts_e.swAllDisplayState, Empty)
    
    If Not IsEmpty(vAppearances) Then
    
        Dim i As Integer
        
        For i = 0 To UBound(vAppearances)
        
            Dim swAppearance As SldWorks.RenderMaterial
            Set swAppearance = vAppearances(i)
            
            Dim vOtherEnts As Variant
            
            If HasFaces(swAppearance, vOtherEnts) Then
            
                swAppearance.RemoveAllEntities
                
                If Not IsEmpty(vOtherEnts) Then
                    Dim j As Integer
                    
                    For j = 0 To UBound(vOtherEnts)
                        If False = swAppearance.AddEntity(vOtherEnts(j)) Then
                            Err.Raise vbError, "", "Failed to restore entity"
                        End If
                    Next
                    
                    Dim matId1 As Long
                    Dim matId2 As Long
                    swAppearance.GetMaterialIds matId1, matId2
                    If False = swModel.Extension.AddDisplayStateSpecificRenderMaterial(swAppearance, swDisplayStateOpts_e.swAllDisplayState, Empty, matId1, matId2) Then
                        Err.Raise vbError, "", "Failed to restore render material"
                    End If
                End If
                
            End If
            
        Next
    
    End If
    
    swModel.EditRebuild3
    
End Sub

Function HasFaces(appearance As SldWorks.RenderMaterial, otherEnts As Variant) As Boolean
    
    HasFaces = False
    
    Dim swOtherEnts() As Object

    Dim vEnts As Variant
    vEnts = appearance.GetEntities
    
    If Not IsEmpty(vEnts) Then
        Dim i As Integer
        For i = 0 To UBound(vEnts)
            If TypeOf vEnts(i) Is SldWorks.Face2 Then
                HasFaces = True
            Else
                If (Not swOtherEnts) = -1 Then
                    ReDim swOtherEnts(0)
                Else
                    ReDim Preserve swOtherEnts(UBound(swOtherEnts) + 1)
                End If
                
                Set swOtherEnts(UBound(swOtherEnts)) = vEnts(i)
            End If
        Next
    End If
    
    If (Not swOtherEnts) = -1 Then
        otherEnts = Empty
    Else
        otherEnts = swOtherEnts
    End If
    
End Function