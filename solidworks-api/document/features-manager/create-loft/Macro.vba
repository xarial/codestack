Dim swApp As SldWorks.SldWorks

Sub main()

    Dim swModel As SldWorks.ModelDoc2
    Dim swSelMgr As SldWorks.SelectionMgr
    
    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc

    Set swSelMgr = swModel.SelectionManager
    
    Dim swFeats() As SldWorks.Feature
    ReDim swFeats(swSelMgr.GetSelectedObjectCount2(-1) - 1)
    
    Dim i As Integer
    
    For i = 1 To swSelMgr.GetSelectedObjectCount2(-1)
        Dim swFeat As SldWorks.Feature
        Set swFeat = swSelMgr.GetSelectedObject6(i, -1)
        Set swFeats(i - 1) = swFeat
    Next
    
    Dim swSelData As SldWorks.SelectData
    Set swSelData = swSelMgr.CreateSelectData
    
    swSelData.Mark = 1
    
    If swModel.Extension.MultiSelect2(swFeats, False, swSelData) <> UBound(swFeats) + 1 Then
        Err.Raise vbError, "", "Failed to selected profiles"
    End If
        
    Const CONSTRAINT_DEFAULT As Integer = 6
    Const THIN_TYPE_ONE_DIR As Integer = 0
    
    swModel.FeatureManager.InsertProtrusionBlend2 False, True, False, 1, CONSTRAINT_DEFAULT, CONSTRAINT_DEFAULT, 1, 1, True, True, False, 0, 0, THIN_TYPE_ONE_DIR, True, True, True, swGuideCurveInfluence_e.swGuideCurveInfluenceNextGuide

End Sub