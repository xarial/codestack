Dim swApp As SldWorks.SldWorks

Sub main()
        
    Set swApp = Application.SldWorks
    
    Dim swModel As SldWorks.ModelDoc2
    
    Dim swRootNode As SldWorks.TreeControlItem
    
    Set swModel = swApp.ActiveDoc
    
    Set swRootNode = swModel.FeatureManager.GetFeatureTreeRootItem2(swFeatMgrPane_e.swFeatMgrPaneBottom)
    
    Dim swBodyFolderNode As SldWorks.TreeControlItem
    Set swBodyFolderNode = TryFindBodyFolderNode(swRootNode)
    
    If Not swBodyFolderNode Is Nothing Then
        swBodyFolderNode.Expanded = True
    Else
        Err.Raise vbError, "", "Failed to find Body Folder"
    End If

End Sub

Function TryFindBodyFolderNode(node As SldWorks.TreeControlItem) As SldWorks.TreeControlItem

    If node.ObjectType = swTreeControlItemType_e.swFeatureManagerItem_Feature Then
        
        Dim swFeat As SldWorks.Feature
        Set swFeat = node.Object
        
        If Not swFeat Is Nothing Then
                        
            Select Case swFeat.GetTypeName2()
                Case "HistoryFolder"
                    Exit Function
                    
                Case "SolidBodyFolder"
                    Set TryFindBodyFolderNode = node
                    Exit Function
            End Select
            
        End If
            
    End If

    Dim swChildNode As SldWorks.TreeControlItem
    
    Set swChildNode = node.GetFirstChild
    
    While Not swChildNode Is Nothing
    
        Dim swBodyFolderNode As SldWorks.TreeControlItem
        Set swBodyFolderNode = TryFindBodyFolderNode(swChildNode)
        
        If Not swBodyFolderNode Is Nothing Then
            Set TryFindBodyFolderNode = swBodyFolderNode
            Exit Function
        End If
        
        Set swChildNode = swChildNode.GetNext()
        
    Wend

End Function