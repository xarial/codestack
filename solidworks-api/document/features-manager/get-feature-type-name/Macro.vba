Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swModel As SldWorks.ModelDoc2
    
    Set swModel = swApp.ActiveDoc
    
    If Not swModel Is Nothing Then
        MsgBox GetTypeNames(swModel.SelectionManager)
    Else
        MsgBox "Please open model"
    End If
    
End Sub

Function GetTypeNames(selMgr As SldWorks.SelectionMgr) As String
    
    Dim typeNames As String
    
    Dim i As Integer
    
    For i = 1 To selMgr.GetSelectedObjectCount2(-1)
        
        On Error Resume Next
        
        Dim swFeat As SldWorks.Feature
        Set swFeat = selMgr.GetSelectedObject6(i, -1)
        
        If Not swFeat Is Nothing Then
            typeNames = typeNames & vbLf & swFeat.Name & ": " & swFeat.GetTypeName() & "; " & swFeat.GetTypeName2
        End If
        
    Next
    
    GetTypeNames = typeNames
    
End Function