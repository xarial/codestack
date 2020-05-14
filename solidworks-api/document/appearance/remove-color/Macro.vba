Const REMOVE_FROM_ALL_CONFIGS As Boolean = True

Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swPart As SldWorks.PartDoc
    
    Set swPart = GetActivePart(swApp)
    
    If Not swPart Is Nothing Then
        
        Dim configOpts As swInConfigurationOpts_e
        configOpts = GetConfigurationOptions(REMOVE_FROM_ALL_CONFIGS)
        
        Dim vBodies As Variant
        vBodies = swPart.GetBodies2(swBodyType_e.swAllBodies, False)
        
        RemoveMaterialPropertiesFromBodies vBodies, True, configOpts
        RemoveMaterialPropertiesFromFeatures swPart.FeatureManager.GetFeatures(False), configOpts
        
    Else
        MsgBox "Please open part document"
    End If
    
End Sub

Sub RemoveMaterialPropertiesFromBodies(bodies As Variant, removeFromFaces As Boolean, configOpts As swInConfigurationOpts_e)
    
    If Not IsEmpty(bodies) Then
        
        Dim i As Integer
        
        For i = 0 To UBound(bodies)
            
            Dim swBody As SldWorks.Body2
            Set swBody = bodies(i)
            
            swBody.RemoveMaterialProperty configOpts, Empty
                        
            If removeFromFaces Then
                Dim vFaces As Variant
                vFaces = swBody.GetFaces()
                RemoveMaterialPropertiesFromFaces vFaces, configOpts
            End If
            
        Next
        
    End If
        
End Sub

Sub RemoveMaterialPropertiesFromFaces(faces As Variant, configOpts As swInConfigurationOpts_e)
    
    Dim i As Integer
    
    If Not IsEmpty(faces) Then
    
        For i = 0 To UBound(faces)
            
            Dim swFace As SldWorks.Face2
            Set swFace = faces(i)
            
            swFace.RemoveMaterialProperty2 configOpts, Empty

        Next
    
    End If
    
End Sub

Sub RemoveMaterialPropertiesFromFeatures(features As Variant, configOpts As swInConfigurationOpts_e)
    
    Dim i As Integer
    
    If Not IsEmpty(features) Then
    
        For i = 0 To UBound(features)
            
            Dim swFeat As SldWorks.Feature
            Set swFeat = features(i)
            
            swFeat.RemoveMaterialProperty2 configOpts, Empty
                
        Next
    
    End If
End Sub

Function GetActivePart(app As SldWorks.SldWorks) As SldWorks.PartDoc
    
    On Error Resume Next
    
    Set GetActivePart = app.ActiveDoc
    
End Function

Function GetConfigurationOptions(allConfigs As Boolean) As swInConfigurationOpts_e
    
    If REMOVE_FROM_ALL_CONFIGS Then
        GetConfigurationOptions = swAllConfiguration
    Else
        GetConfigurationOptions = swThisConfiguration
    End If
    
End Function