Const OUT_COORD_SYSTEM_NAME As String = "Export CS"
Const OUT_EXTENSION As String = ".x_t"

Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swModel As SldWorks.ModelDoc2
    Set swModel = swApp.ActiveDoc
    
    If Not swModel Is Nothing Then
        
        Dim outFilePath As String
        outFilePath = ComposeOutputFilePath(swModel, OUT_EXTENSION)
        
        Export swModel, outFilePath, OUT_COORD_SYSTEM_NAME
        
    Else
        Err.Raise vbError, "", "Please open the model"
    End If
    
End Sub

Function ComposeOutputFilePath(model As SldWorks.ModelDoc2, ext As String) As String
    
    Dim path As String
    path = model.GetPathName
    
    If path <> "" Then
        ComposeOutputFilePath = GetDirectory(path) & GetFileNameWithoutExtension(path) + ext
    Else
        Err.Raise vbError, "", "Model is not saved on disk"
    End If
    
End Function

Function GetDirectory(path As String)
    GetDirectory = Left(path, InStrRev(path, "\"))
End Function

Function GetFileNameWithoutExtension(path As String) As String
    GetFileNameWithoutExtension = Mid(path, InStrRev(path, "\") + 1, InStrRev(path, ".") - InStrRev(path, "\") - 1)
End Function

Sub Export(model As SldWorks.ModelDoc2, path As String, coordSysName As String)
    
    Dim curOutCoordSys As String
    curOutCoordSys = swApp.GetUserPreferenceStringValue(swUserPreferenceStringValue_e.swExportOutputCoordinateSystem)

    swApp.SetUserPreferenceStringValue swUserPreferenceStringValue_e.swExportOutputCoordinateSystem, coordSysName
    
    Dim errors As Long
    Dim warnings As Long
    
    If False = model.Extension.SaveAs(path, swSaveAsVersion_e.swSaveAsCurrentVersion, swSaveAsOptions_e.swSaveAsOptions_Silent, Nothing, errors, warnings) Then
        Err.Raise vbError, "", "Failed to export file. Error code: " & errors
    End If
    
    swApp.SetUserPreferenceStringValue swUserPreferenceStringValue_e.swExportOutputCoordinateSystem, curOutCoordSys
        
End Sub