Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
        
    Dim swAssy As SldWorks.AssemblyDoc
    Set swAssy = TryGetActiveAssembly
    
    If Not swAssy Is Nothing Then
        swApp.RunCommand swCommands_VisualizationTool, ""
    Else
        MsgBox "Please open assembly"
    End If
    
End Sub

Function TryGetActiveAssembly() As SldWorks.AssemblyDoc
    
    Dim swModel As SldWorks.ModelDoc2
    
    Set swModel = swApp.ActiveDoc
    
    If Not swModel Is Nothing Then
        
        If swModel.GetType() = swDocumentTypes_e.swDocASSEMBLY Then
            Set TryGetActiveAssembly = swApp.ActiveDoc
        End If
        
    End If
    
End Function