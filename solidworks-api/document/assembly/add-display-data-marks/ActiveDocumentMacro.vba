Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swModel As SldWorks.ModelDoc2
    
    Set swModel = swApp.ActiveDoc
    
    If Not swModel Is Nothing Then
            
        If swModel.GetType() = swDocumentTypes_e.swDocASSEMBLY Or swModel.GetType() = swDocumentTypes_e.swDocPART Then
            
            Dim vConfNames As Variant
            vConfNames = swModel.GetConfigurationNames
            
            Dim i As Integer
            
            For i = 0 To UBound(vConfNames)
                Dim swConf As SldWorks.Configuration
                Set swConf = swModel.GetConfigurationByName(CStr(vConfNames(i)))
                swConf.LargeDesignReviewMark = True
            Next
            
            swModel.ForceRebuild3 False
            
        Else
            Err.Raise vbError, "", "Only assemblies and parts are supported"
        End If
        
    Else
        Err.Raise vbError, "", "No files opened"
    End If
    
End Sub