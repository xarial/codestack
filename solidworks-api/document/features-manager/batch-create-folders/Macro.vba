Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swModel As SldWorks.ModelDoc2
    
    Set swModel = swApp.ActiveDoc
    
    If Not swModel Is Nothing Then
    
        Dim foldersCount As Integer
        Dim folderNamePrefix As String
        
        foldersCount = CInt(InputBox("Specify the number of folders to create", "Batch Folder Creator", "5"))
        folderNamePrefix = InputBox("Specify the prefix name of the folder", "Batch Folder Creator", "MyFolder")
        
        Dim swAnchorFeat As SldWorks.Feature
        Set swAnchorFeat = swModel.Extension.GetLastFeatureAdded
        
        Dim swFeatMgr As SldWorks.FeatureManager
        Set swFeatMgr = swModel.FeatureManager
        
        Dim i As Integer
        
        Dim nextIndex As Integer
        nextIndex = 0
        
        For i = 1 To foldersCount
            
            swAnchorFeat.Select2 False, -1
            
            Dim swFolderFeat As SldWorks.Feature
            Set swFolderFeat = swFeatMgr.InsertFeatureTreeFolder2(swFeatureTreeFolderType_e.swFeatureTreeFolder_EmptyBefore)
            
            If swFolderFeat Is Nothing Then
                Err.Raise vbError, "", "Failed to create a folder, make sure there there is at least one feature in the model"
            End If
            
            Dim folderName As String
            
            Do
                nextIndex = nextIndex + 1
                folderName = folderNamePrefix & nextIndex
            Loop While False <> swFeatMgr.IsNameUsed(swNameType_e.swFeatureName, folderName)
            
            swFolderFeat.Name = folderName
            
            swModel.Extension.ReorderFeature swFolderFeat.Name, "", swMoveLocation_e.swMoveToEnd
            
        Next
        
    Else
        Err.Raise vbError, "", "No model opened"
    End If
    
End Sub