Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swModel As SldWorks.ModelDoc2

    Set swModel = swApp.ActiveDoc
    
    If Not swModel Is Nothing Then
        
        Dim vObjects As Variant
        vObjects = GetAllSelectedObjects(swModel)
        
        swModel.ClearSelection2 True
        
        Stop
        
        swModel.Extension.MultiSelect2 vObjects, False, Nothing
        
    Else
        MsgBox "Please open the document"
    End If
    
End Sub

Function GetAllSelectedObjects(model As SldWorks.ModelDoc2) As Variant
    
    Dim swSelMgr As SldWorks.SelectionMgr
    Dim swObjects() As Object
    
    Set swSelMgr = model.SelectionManager
    
    Dim i As Integer
    
    For i = 1 To swSelMgr.GetSelectedObjectCount2(-1)
        
        Dim swObj As Object
        Set swObj = swSelMgr.GetSelectedObject6(i, -1)
        
        ReDim Preserve swObjects(i - 1)
        Set swObjects(i - 1) = swObj
    Next
    
    GetAllSelectedObjects = swObjects
    
End Function