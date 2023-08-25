Const UNUBSORBED_ONLY As Boolean = True

Dim swApp As SldWorks.SldWorks
Dim swModel As SldWorks.ModelDoc2
Dim swPart As SldWorks.PartDoc

Sub main()

    Set swApp = Application.SldWorks

    Set swModel = swApp.ActiveDoc
        
    Set swPart = swModel
        
    Dim vFeats As Variant
    
    vFeats = CollectSelectedSketches(swModel)
    
    If IsEmpty(vFeats) Then
        vFeats = CollectAllSketchFeatures(swModel.FirstFeature)
    End If
    
    If Not IsEmpty(vFeats) Then
        
        Dim i As Integer
        
        For i = 0 To UBound(vFeats)
            
            Dim swFeat As SldWorks.Feature
            Set swFeat = vFeats(i)
                        
            If False <> swFeat.Select2(False, -1) Then
                swPart.SetLineColor RGB(CInt(255 * Rnd()), CInt(255 * Rnd()), CInt(255 * Rnd()))
            Else
                Err.Raise vbError, "", "Failed to select " & swFeat.Name
            End If
            
        Next
        
    End If
    
    swModel.ClearSelection2 True

End Sub

Function IsAbsorbed(feat As SldWorks.Feature) As Boolean
    
    Dim vFeatChildren As Variant
    vFeatChildren = feat.GetChildren()
    
    IsAbsorbed = Not IsEmpty(vFeatChildren)
    
End Function

Function CollectSelectedSketches(model As SldWorks.ModelDoc2) As Variant
    
    Dim swFeats() As SldWorks.Feature

    Dim swSelMgr As SldWorks.SelectionMgr
    
    Set swSelMgr = model.SelectionManager
    
    Dim i As Integer
    
    For i = 1 To swSelMgr.GetSelectedObjectCount2(-1)
        
        If swSelMgr.GetSelectedObjectType3(i, -1) = swSelectType_e.swSelSKETCHES Then
            
            If (Not swFeats) = -1 Then
                ReDim swFeats(0)
            Else
                ReDim Preserve swFeats(UBound(swFeats) + 1)
            End If
            
            Set swFeats(UBound(swFeats)) = swSelMgr.GetSelectedObject6(i, -1)
            
        End If
        
    Next
    
    If (Not swFeats) = -1 Then
        CollectSelectedSketches = Empty
    Else
        CollectSelectedSketches = swFeats
    End If

End Function

Function CollectAllSketchFeatures(firstFeat As SldWorks.Feature) As Variant
    
    Const SKETCH_FEAT_TYPE_NAME As String = "ProfileFeature"
    Const SKETCH_3D_FEAT_TYPE_NAME As String = "3DProfileFeature"

    Dim swFeats() As SldWorks.Feature

    Dim swFeat As SldWorks.Feature
    Set swFeat = firstFeat
    
    While Not swFeat Is Nothing
    
        If swFeat.GetTypeName2 = SKETCH_FEAT_TYPE_NAME Or _
            swFeat.GetTypeName2 = SKETCH_3D_FEAT_TYPE_NAME Then
            
            If Not UNUBSORBED_ONLY Or Not IsAbsorbed(swFeat) Then
            
                If (Not swFeats) = -1 Then
                    ReDim swFeats(0)
                Else
                    ReDim Preserve swFeats(UBound(swFeats) + 1)
                End If
                
                Set swFeats(UBound(swFeats)) = swFeat
            
            End If
            
        End If
        
        Set swFeat = swFeat.GetNextFeature
        
    Wend
    
    If (Not swFeats) = -1 Then
        CollectAllSketchFeatures = Empty
    Else
        CollectAllSketchFeatures = swFeats
    End If
    
End Function