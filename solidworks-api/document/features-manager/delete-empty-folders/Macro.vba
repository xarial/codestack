Dim swApp As SldWorks.SldWorks

Sub main()
       
    Set swApp = Application.SldWorks
    
    Dim swModel As SldWorks.ModelDoc2
    
    Set swModel = swApp.ActiveDoc
    
    Dim swFeat As SldWorks.Feature
    
    Set swFeat = swModel.FirstFeature
    
    swModel.ClearSelection2 True
    
    While Not swFeat Is Nothing
        
        If IsFolder(swFeat) Then
            
            If Not FolderHasFeatures(swFeat.GetSpecificFeature2()) Then
                swFeat.Select2 True, -1
            End If
            
        End If
        
        Dim swSubFeat As SldWorks.Feature
        Set swSubFeat = swFeat.GetFirstSubFeature
        
        While Not swSubFeat Is Nothing
            
            If IsFolder(swSubFeat) Then
                
                If Not FolderHasFeatures(swSubFeat.GetSpecificFeature2()) Then
                    swSubFeat.Select2 True, -1
                End If
                
            End If
            
            Set swSubFeat = swSubFeat.GetNextSubFeature
        Wend
        
        Set swFeat = swFeat.GetNextFeature
        
    Wend
    
    swModel.Extension.DeleteSelection2 swDeleteSelectionOptions_e.swDelete_Absorbed
    
End Sub

Function IsFolder(feat As SldWorks.Feature) As Boolean
    Const END_FOLDER_TAG As String = "___EndTag___"
    IsFolder = feat.GetTypeName2() = "FtrFolder" And Right(feat.Name, Len(END_FOLDER_TAG)) <> END_FOLDER_TAG
End Function

Function FolderHasFeatures(folder As SldWorks.FeatureFolder) As Boolean
    
    If folder.GetFeatureCount() = 0 Then
        FolderHasFeatures = False
    Else
        
        Dim vFeats As Variant
        
        vFeats = folder.GetFeatures()
        Dim i As Integer
        
        For i = 0 To UBound(vFeats)
            
            Dim swFeat As SldWorks.Feature
            Set swFeat = vFeats(i)
            
            If IsFolder(swFeat) Then
                If FolderHasFeatures(swFeat.GetSpecificFeature2()) Then
                    FolderHasFeatures = True
                    Exit Function
                End If
            Else
                FolderHasFeatures = True
                Exit Function
            End If
            
        Next
        
        FolderHasFeatures = False
        
    End If
    
End Function