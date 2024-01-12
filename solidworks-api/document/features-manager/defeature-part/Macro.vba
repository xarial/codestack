Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
try_:
    
    On Error GoTo catch_
    
    Dim swPart As SldWorks.PartDoc
    
    Set swPart = swApp.ActiveDoc
    
    If Not swPart Is Nothing Then
        
        Dim vUserFeats As Variant
        vUserFeats = GetAllTopLevelUserFeatures(swPart)
        
        If Not IsEmpty(vUserFeats) Then
            CreateFeaturesForBodies swPart
            DeleteFeatures swPart, vUserFeats
        Else
            Err.Raise vbError, "", "No features in the model"
        End If
        
    Else
        MsgBox "Please open part document"
    End If
    
    GoTo finally_
    
catch_:
    MsgBox Err.Description, vbCritical
finally_:
    
End Sub

Sub CreateFeaturesForBodies(part As SldWorks.PartDoc)
    
    Dim vBodies As Variant
    
    vBodies = part.GetBodies2(swBodyType_e.swAllBodies, False)
    
    If Not IsEmpty(vBodies) Then
                
        Dim i As Integer
        
        For i = 0 To UBound(vBodies)
            
            Dim swBody As SldWorks.Body2
            Set swBody = vBodies(i)
            Set swBodyCopy = swBody.Copy()
                        
            Dim swFeat As SldWorks.Feature
        
            Set swFeat = part.CreateFeatureFromBody3(swBodyCopy, False, swCreateFeatureBodyOpts_e.swCreateFeatureBodySimplify)
            
            If Not swFeat Is Nothing Then
                
                Dim swFace As SldWorks.Face2
                Set swFace = swFeat.GetFaces()(0)
                
                Dim swReplacedBody As SldWorks.Body2
                Set swReplacedBody = swFace.GetBody
                
                swReplacedBody.HideBody False = swBody.Visible
                
            Else
                Err.Raise vbError, "", "Failed to create feature for a body " & swBody.Name
            End If
                        
        Next
    
    Else
        
        Err.Raise vbError, "", "No bodies found"
        
    End If
    
End Sub

Sub DeleteFeatures(model As SldWorks.ModelDoc2, feats As Variant)
    
    If model.Extension.MultiSelect2(feats, False, Nothing) = UBound(feats) + 1 Then
        model.Extension.DeleteSelection2 swDeleteSelectionOptions_e.swDelete_Children + swDeleteSelectionOptions_e.swDelete_Absorbed
    Else
        Err.Raise vbError, "", "Failed to select user features"
    End If
            
End Sub

Function GetAllTopLevelUserFeatures(model As SldWorks.ModelDoc2) As Variant
    
    Dim swUserFeats() As SldWorks.Feature
    
    Dim swFeat As SldWorks.Feature
    
    Set swFeat = model.FirstFeature
    
    Dim isUserFeat As Boolean
    isUserFeat = False
    
    While Not swFeat Is Nothing
        
        If isUserFeat Then

            If (Not swUserFeats) = -1 Then
                ReDim swUserFeats(0)
            Else
                ReDim Preserve swUserFeats(UBound(swUserFeats) + 1)
            End If
            
            Set swUserFeats(UBound(swUserFeats)) = swFeat
        
        Else
            If swFeat.GetTypeName2() = "OriginProfileFeature" Then
                isUserFeat = True
            End If
        End If
        
        Set swFeat = swFeat.GetNextFeature
        
    Wend
    
    If (Not swUserFeats) = -1 Then
        GetAllTopLevelUserFeatures = Empty
    Else
        GetAllTopLevelUserFeatures = swUserFeats
    End If
    
End Function