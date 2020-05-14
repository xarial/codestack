Const OFFSET_SYMBOL = " "

Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swModel As SldWorks.ModelDoc2
    
    Set swModel = swApp.ActiveDoc

    If Not swModel Is Nothing Then
    
        Dim swFeatMgr As SldWorks.FeatureManager
        
        Set swFeatMgr = swModel.FeatureManager
        
        Dim swRootFeatNode As SldWorks.TreeControlItem
        
        Set swRootFeatNode = swFeatMgr.GetFeatureTreeRootItem2(swFeatMgrPane_e.swFeatMgrPaneBottom)
        
        If Not swRootFeatNode Is Nothing Then
            TraverseFeatureNode swRootFeatNode, ""
        End If
        
    Else
        MsgBox "Please open the model"
    End If
End Sub

Sub TraverseFeatureNode(featNode As SldWorks.TreeControlItem, offset As String)
    
    Debug.Print offset & featNode.Text
    
    Dim swChildFeatNode As SldWorks.TreeControlItem
    
    Set swChildFeatNode = featNode.GetFirstChild()
    
    While Not swChildFeatNode Is Nothing
        TraverseFeatureNode swChildFeatNode, offset + OFFSET_SYMBOL
        Set swChildFeatNode = swChildFeatNode.GetNext
    Wend
    
End Sub