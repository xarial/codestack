Dim swController As Controller

Function swmRebuild(varApp As Variant, varDoc As Variant, varFeat As Variant) As Variant
    
    Dim swApp As SldWorks.SldWorks
    Dim swModel As SldWorks.ModelDoc2
    Dim swFeat As SldWorks.Feature
    Dim swMacroFeatData As SldWorks.MacroFeatureData
    
    Set swApp = varApp
    Set swModel = varDoc
    Set swFeat = varFeat
    
    Set swMacroFeatData = swFeat.GetDefinition
            
    Dim vSketches As Variant
    Dim vDepths As Variant
    
    swMacroFeatData.GetSelections3 vSketches, Empty, Empty, Empty, Empty
    swMacroFeatData.GetParameters Empty, Empty, vDepths
    
    Dim vBodies As Variant
    vBodies = Geometry.CreateBodiesFromSketches(vSketches, vDepths)
    
    Dim i As Integer
    
    For i = 0 To UBound(vBodies)
        Dim swBody As SldWorks.Body2
        Set swBody = vBodies(i)
        AssignUserIds swBody, swMacroFeatData
    Next
    
    swMacroFeatData.EnableMultiBodyConsume = True
    
    swmRebuild = vBodies
    
End Function

Sub AssignUserIds(body As SldWorks.Body2, featData As SldWorks.MacroFeatureData)
    
    Dim vFaces As Variant
    Dim vEdges As Variant
    Dim i As Integer
    
    featData.GetEntitiesNeedUserId body, vFaces, vEdges
    
    If Not IsEmpty(vFaces) Then
        For i = 0 To UBound(vFaces)
            Dim swFace As SldWorks.Face2
            Set swFace = vFaces(i)
            featData.SetFaceUserId swFace, 0, i
        Next
    End If
    
    If Not IsEmpty(vEdges) Then
        For i = 0 To UBound(vEdges)
            Dim swEdge As SldWorks.Edge
            Set swEdge = vEdges(i)
            featData.SetEdgeUserId swEdge, 0, i
        Next
    End If
    
End Sub

Function swmEditDefinition(varApp As Variant, varDoc As Variant, varFeat As Variant) As Variant
    
    If swController Is Nothing Then
        Set swController = New Controller
    End If
    
    Dim swFeat As SldWorks.Feature
    Set swFeat = varFeat
    
    swController.EditExtrude swFeat
    
    swmEditDefinition = True
    
End Function

Function swmSecurity(varApp As Variant, varDoc As Variant, varFeat As Variant) As Variant
    swmSecurity = swMacroFeatureSecurityOptions_e.swMacroFeatureSecurityByDefault
End Function