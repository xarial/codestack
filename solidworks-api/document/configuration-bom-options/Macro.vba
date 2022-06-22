Const ALL_CONFIGS As Boolean = True
Const PART_NUMBER_SRC As Integer = swBOMPartNumberSource_e.swBOMPartNumber_ConfigurationName
Const CHILD_COMPS_DISP As Integer = swChildComponentInBOMOption_e.swChildComponent_Promote

Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swModel As SldWorks.ModelDoc2
    
    Set swModel = swApp.ActiveDoc
    
    If swModel Is Nothing Then
        Err.Raise vbError, "", "Open document"
    End If
    
    If swModel.GetType() = swDocumentTypes_e.swDocDRAWING Then
        Err.Raise vbError, "", "Drawings are not supported"
    End If
    
    If ALL_CONFIGS Then
        
        Dim vConfNames As Variant
        
        vConfNames = swModel.GetConfigurationNames
        Dim i As Integer
        
        For i = 0 To UBound(vConfNames)
            Dim swConf As SldWorks.Configuration
            Set swConf = swModel.GetConfigurationByName(CStr(vConfNames(i)))
            SetConfigurationBomOptions swConf
        Next
        
    Else
        SetConfigurationBomOptions swModel.ConfigurationManager.ActiveConfiguration
    End If
    
End Sub

Sub SetConfigurationBomOptions(config As SldWorks.Configuration)
    config.ChildComponentDisplayInBOM = CHILD_COMPS_DISP
    config.BOMPartNoSource = PART_NUMBER_SRC
End Sub