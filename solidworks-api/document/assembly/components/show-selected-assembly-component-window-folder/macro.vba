Dim swApp As SldWorks.SldWorks
Dim swModel As SldWorks.ModelDoc2
Dim swSelMgr As SldWorks.SelectionMgr
Dim swComp As SldWorks.Component2

Sub main()

    On Error Resume Next
    
    Set swApp = Application.SldWorks
    
    Set swModel = swApp.ActiveDoc
    
    If Not swModel Is Nothing Then
    
        Set swSelMgr = swModel.SelectionManager
        
        Set swComp = swSelMgr.GetSelectedObjectsComponent3(1, -1)

        Dim path As String
        
        If Not swComp Is Nothing Then
            path = swComp.GetPathName
        Else
            path = swModel.GetPathName
        End If
        
        If path <> "" Then
            Shell "explorer.exe /select, " & """" & path & """", vbMaximizedFocus
        Else
            MsgBox "Model is not saved"
        End If
    
    Else
        MsgBox "Please open assembly document and select the component"
    End If
    
End Sub