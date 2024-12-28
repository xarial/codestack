Const SHADED_DRAFT_QUALITY_RESOLUTION As Double = 0.5 '0 to 1
Const WIREFRAME_HIGHT_QUALITY_RESOLUTION As Double = 0.5 '0 to 1

Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swModel As SldWorks.ModelDoc2
    
    Set swModel = swApp.ActiveDoc
    
    Const QUALITY_MIN As Integer = 0
    Const QUALITY_MAX As Integer = 106
    
    Dim qualVal As Integer
    qualVal = CInt(QUALITY_MIN + (QUALITY_MAX - QUALITY_MIN) * SHADED_DRAFT_QUALITY_RESOLUTION)
    
    swModel.SetTessellationQuality qualVal
    
    Const WIREFRAME_HIGHT_QUALITY_RESOLUTION_MIN As Integer = 1
    Const WIREFRAME_HIGHT_QUALITY_RESOLUTION_MAX As Integer = 100
    
    Dim wireframeResVal As Integer
    wireframeResVal = CInt(WIREFRAME_HIGHT_QUALITY_RESOLUTION_MIN + (WIREFRAME_HIGHT_QUALITY_RESOLUTION_MAX - WIREFRAME_HIGHT_QUALITY_RESOLUTION_MIN) * WIREFRAME_HIGHT_QUALITY_RESOLUTION)
    
    If False = swModel.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swImageQualityWireframeValue, swUserPreferenceOption_e.swDetailingNoOptionSpecified, wireframeResVal) Then
        Err.Raise vbError, "", "Failed to set the image quality wireframe resolution"
    End If
    
    swModel.SetSaveFlag
    
End Sub