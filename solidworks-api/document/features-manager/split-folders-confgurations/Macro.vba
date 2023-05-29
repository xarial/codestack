Const CREATE_DERIVED_CONFS As Boolean = True

Const FOLDER_END_TAG As String = "___EndTag___"

Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swModel As SldWorks.ModelDoc2
    
    Set swModel = swApp.ActiveDoc
    
    If Not swModel Is Nothing Then
        
        Dim vFeatFolders As Variant
        Dim vAllFeatFolders As Variant
        
        Dim swSelMgr As SldWorks.SelectionMgr
        Set swSelMgr = swModel.SelectionManager
        
        vAllFeatFolders = GetAllFeatureFolders(swModel)
        
        If swSelMgr.GetSelectedObjectCount2(-1) = 0 Then
            vFeatFolders = vAllFeatFolders
        Else
            vFeatFolders = GetSelectedFeatureFolders(swModel)
        End If
        
        If Not IsEmpty(vFeatFolders) Then
            
            Dim activeConfName As String
            activeConfName = swModel.ConfigurationManager.ActiveConfiguration.Name
            
            Dim i As Integer
            
            For i = 0 To UBound(vFeatFolders)
                Dim swFeatFolder As SldWorks.Feature
                Set swFeatFolder = vFeatFolders(i)
                CreateConfigurationForFolder swModel, swFeatFolder, vAllFeatFolders, IIf(CREATE_DERIVED_CONFS, activeConfName, "")
            Next
            
        End If
                
    Else
        Err.Raise vbError, "", "No active document"
    End If
    
End Sub

Function GetAllFeatureFolders(model As SldWorks.ModelDoc2) As Variant
    
    Dim swFeatFolders() As SldWorks.Feature
    
    Dim swFeat As SldWorks.Feature
    Set swFeat = model.FirstFeature
    
    While Not swFeat Is Nothing
        
        If swFeat.GetTypeName2() = "FtrFolder" And InStr(LCase(swFeat.Name), LCase(FOLDER_END_TAG)) = 0 Then

            If (Not swFeatFolders) = -1 Then
                ReDim swFeatFolders(0)
            Else
                ReDim Preserve swFeatFolders(UBound(swFeatFolders) + 1)
            End If
            
            Set swFeatFolders(UBound(swFeatFolders)) = swFeat
            
        End If
        
        Set swFeat = swFeat.GetNextFeature
        
    Wend
    
    
    If (Not swFeatFolders) = -1 Then
        GetAllFeatureFolders = Empty
    Else
        GetAllFeatureFolders = swFeatFolders
    End If
        
End Function

Function GetSelectedFeatureFolders(model As SldWorks.ModelDoc2) As Variant
    
    Dim swSelMgr As SldWorks.SelectionMgr
    Set swSelMgr = model.SelectionManager

    Dim swFeatFolders() As SldWorks.Feature
    
    Dim i As Integer
    
    For i = 1 To swSelMgr.GetSelectedObjectCount2(-1)
        
        If swSelMgr.GetSelectedObjectType3(i, -1) = swSelectType_e.swSelFTRFOLDER Then
        
            Dim swFeat As SldWorks.Feature
            Set swFeat = swSelMgr.GetSelectedObject6(i, -1)
            
            If (Not swFeatFolders) = -1 Then
                ReDim swFeatFolders(0)
            Else
                ReDim Preserve swFeatFolders(UBound(swFeatFolders) + 1)
            End If
            
            Set swFeatFolders(UBound(swFeatFolders)) = swFeat
        End If
    
    Next
        
    If (Not swFeatFolders) = -1 Then
        GetSelectedFeatureFolders = Empty
    Else
        GetSelectedFeatureFolders = swFeatFolders
    End If
    
End Function

Sub CreateConfigurationForFolder(model As SldWorks.ModelDoc2, folderFeat As SldWorks.Feature, allFeatFolders As Variant, parentConfName As String)
    
    Dim swFolderConf As SldWorks.Configuration
    Set swFolderConf = model.ConfigurationManager.AddConfiguration2(folderFeat.Name, "", "", swConfigurationOptions2_e.swConfigOption_DontActivate Or swConfigurationOptions2_e.swConfigOption_SuppressByDefault, parentConfName, "", False)
    
    If swFolderConf Is Nothing Then
        Err.Raise vbError, "", "Failed to create configuration for " & folderFeat.Name
    End If
    
    Dim i As Integer
    
    For i = 0 To UBound(allFeatFolders)
        
        Dim swOtherFeatFolder As SldWorks.Feature
        Set swOtherFeatFolder = allFeatFolders(i)
        
        If swApp.IsSame(folderFeat, swOtherFeatFolder) <> swObjectEquality.swObjectSame Then
        
            Dim targetConf(0) As String
            targetConf(0) = swFolderConf.Name
            
            If False = swOtherFeatFolder.SetSuppression2(swFeatureSuppressionAction_e.swSuppressFeature, swInConfigurationOpts_e.swSpecifyConfiguration, targetConf) Then
                Err.Raise vbError, "", "Failed to configure the suppression of the folder feature for " & swOtherFeatFolder.Name & " in " & swFolderConf.Name
            End If
            
        End If
        
    Next
    
End Sub