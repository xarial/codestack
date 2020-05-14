Enum IfcFormat_e
    Ifc2x3 = 23
    Ifc4 = 4
End Enum

Const OUT_FILE_PATH As String = "C:\Engine.ifc"

Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swModel As SldWorks.ModelDoc2
    Set swModel = swApp.ActiveDoc
    
    If Not swModel Is Nothing Then
        
        ExportIfc swModel, OUT_FILE_PATH, IfcFormat_e.Ifc4
        
    Else
        MsgBox "Please open the model"
    End If
    
End Sub

Sub ExportIfc(model As SldWorks.ModelDoc2, path As String, format As IfcFormat_e)
    
    Dim curIfcFormat As Integer
    curIfcFormat = swApp.GetUserPreferenceIntegerValue(swUserPreferenceIntegerValue_e.swSaveIFCFormat)

    swApp.SetUserPreferenceIntegerValue swUserPreferenceIntegerValue_e.swSaveIFCFormat, format
    
    Dim errors As Long
    Dim warnings As Long
    
    If False = model.Extension.SaveAs(path, swSaveAsVersion_e.swSaveAsCurrentVersion, swSaveAsOptions_e.swSaveAsOptions_Silent, Nothing, errors, warnings) Then
        Err.Raise vbError, "", "Failed to export file. Error code: " & errors
    End If
    
    swApp.SetUserPreferenceIntegerValue swUserPreferenceIntegerValue_e.swSaveIFCFormat, curIfcFormat
        
End Sub