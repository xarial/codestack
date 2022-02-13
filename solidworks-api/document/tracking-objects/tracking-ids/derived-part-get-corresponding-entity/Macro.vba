Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swSrcModel As SldWorks.ModelDoc2
    
    Set swSrcModel = swApp.ActiveDoc
    
    If swSrcModel.GetType() <> swDocumentTypes_e.swDocPART Then
        Err.Raise vbError, "", "Only parts are supported"
    End If
    
    Dim trackDefId As Integer
    trackDefId = TrackSelectedEntities(swSrcModel)
    
    Stop
    
    Dim swTargModel As SldWorks.ModelDoc2
    Set swTargModel = swApp.ActiveDoc
    
    Dim swTargPart As SldWorks.PartDoc
    Set swTargPart = swTargModel
    
    Dim swDerPartFeat As SldWorks.Feature
    
    Set swDerPartFeat = swTargPart.InsertPart3(swSrcModel.GetPathName(), swInsertPartOptions_e.swInsertPartImportSolids, swSrcModel.ConfigurationManager.ActiveConfiguration.Name)
    
    Dim vTrackedEnts As Variant
    vTrackedEnts = GetTrackedEntitites(swTargModel, swDerPartFeat, trackDefId)
    
    If Not IsEmpty(vTrackedEnts) Then
        swTargModel.Extension.MultiSelect2 vTrackedEnts, False, Nothing
    Else
        Err.Raise vbError, "", "No tracked entities found"
    End If
    
End Sub

Function TrackSelectedEntities(model As SldWorks.ModelDoc2) As Integer
    
    Dim trackDefId As Integer
    
    trackDefId = swApp.RegisterTrackingDefinition("_DerivedPartTrack_")
    
    Dim i As Integer
    
    For i = 1 To model.SelectionManager.GetSelectedObjectCount2(-1)
            
        Select Case model.SelectionManager.GetSelectedObjectType3(i, -1)
            Case swSelectType_e.swSelFACES
                Dim swFace As SldWorks.Face2
                Set swFace = model.SelectionManager.GetSelectedObject6(i, -1)
                If swFace.SetTrackingID(trackDefId, i) <> swTrackingIDError_e.swTrackingIDError_NoError Then
                    Err.Raise vbError, "", "Failed to track face"
                End If
            Case swSelectType_e.swSelEDGES
                Dim swEdge As SldWorks.Edge
                Set swEdge = model.SelectionManager.GetSelectedObject6(i, -1)
                If swEdge.SetTrackingID(trackDefId, i) <> swTrackingIDError_e.swTrackingIDError_NoError Then
                    Err.Raise vbError, "", "Failed to track edge"
                End If
            Case swSelectType_e.swSelVERTICES
                Dim swVertex As SldWorks.Vertex
                Set swVertex = model.SelectionManager.GetSelectedObject6(i, -1)
                If swVertex.SetTrackingID(trackDefId, i) <> swTrackingIDError_e.swTrackingIDError_NoError Then
                    Err.Raise vbError, "", "Failed to track vertex"
                End If
            Case Else
                Err.Raise vbError, "", "Only faces, edges and vertices are supported"
        End Select
        
    Next
    
    TrackSelectedEntities = trackDefId
    
End Function

Function GetTrackedEntitites(model As SldWorks.ModelDoc2, derFeatPart As SldWorks.Feature, trackDefId As Integer) As Variant

    Dim isInit As Boolean
    isInit = False
    Dim swEnts() As SldWorks.Entity
    
    Dim searchTypes(2) As Integer
    searchTypes(0) = swTopoEntity_e.swTopoFace
    searchTypes(1) = swTopoEntity_e.swTopoEdge
    searchTypes(2) = swTopoEntity_e.swTopoVertex
    
    Dim vBodies As Variant
    vBodies = GetFeatureBodies(derFeatPart)
    
    Dim i As Integer
    
    For i = 0 To UBound(vBodies)
    
        Dim vTrackedEnts As Variant
        Dim swBody As SldWorks.Body2
        Set swBody = vBodies(i)
        
        vTrackedEnts = model.Extension.FindTrackedObjects(trackDefId, swBody, searchTypes, Empty)
        
        If Not IsEmpty(vTrackedEnts) Then
            If Not isInit Then
                isInit = True
                ReDim swEnts(UBound(vTrackedEnts))
            Else
                ReDim Preserve swEnts(UBound(swEnts) + UBound(vTrackedEnts) + 1)
            End If
            
            Dim j As Integer
            
            For j = 0 To UBound(vTrackedEnts)
                Dim swEnt As SldWorks.Entity
                Set swEnt = vTrackedEnts(j)
                Set swEnts(UBound(swEnts) - UBound(vTrackedEnts) + j) = swEnt
            Next
            
        End If
    
    Next

    If isInit Then
        GetTrackedEntitites = swEnts
    Else
        GetTrackedEntitites = Empty
    End If

End Function

Function GetFeatureBodies(feat As SldWorks.Feature) As Variant
    
    Dim isInit As Boolean
    isInit = False
    
    Dim swBodies() As SldWorks.Body2

    Dim i As Integer
    
    Dim vFaces As Variant
    
    vFaces = feat.GetFaces
    
    For i = 0 To UBound(vFaces)
                
        Dim swFace As SldWorks.Face2
    
        Set swFace = vFaces(i)
        
        Dim swBody As SldWorks.Body2
        
        Set swBody = swFace.GetBody
        
            If Not isInit Then
                ReDim swBodies(0)
                Set swBodies(0) = swBody
                isInit = True
            Else
                If Not Contains(swBodies, swBody) Then
                    ReDim Preserve swBodies(UBound(swBodies) + 1)
                    Set swBodies(UBound(swBodies)) = swBody
                End If
            End If
    
    Next

    If isInit Then
        GetFeatureBodies = swBodies
    Else
        GetFeatureBodies = Empty
    End If

End Function

Function Contains(vArr As Variant, item As Object) As Boolean
    
    Dim i As Integer
    
    For i = 0 To UBound(vArr)
        If vArr(i) Is item Then
            Contains = True
            Exit Function
        End If
    Next
    
    Contains = False
    
End Function