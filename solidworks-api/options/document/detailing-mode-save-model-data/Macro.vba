Const ENABLE As Boolean = True

Const swCommands_Save As Long = 2

Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swModel As SldWorks.ModelDoc2
    
    Set swModel = swApp.ActiveDoc
    
    If Not swModel Is Nothing Then
        
        If swModel.GetType() = swDocumentTypes_e.swDocDRAWING Then
            Dim saveModelDataOpt As Boolean
            Dim includeStandardView As Boolean
            
            saveModelDataOpt = swModel.Extension.GetUserPreferenceToggle(swUserPreferenceToggle_e.swDetailingModeSaveModelData, swUserPreferenceOption_e.swDetailingNoOptionSpecified)
            includeStandardView = swModel.Extension.GetUserPreferenceToggle(swUserPreferenceToggle_e.swDetailingModeIncludeStandardViewsInViewPalette, swUserPreferenceOption_e.swDetailingNoOptionSpecified)
            
            swModel.Extension.SetUserPreferenceToggle swUserPreferenceToggle_e.swDetailingModeSaveModelData, swUserPreferenceOption_e.swDetailingNoOptionSpecified, ENABLE
            swModel.Extension.SetUserPreferenceToggle swUserPreferenceToggle_e.swDetailingModeIncludeStandardViewsInViewPalette, swUserPreferenceOption_e.swDetailingNoOptionSpecified, ENABLE
            
            swApp.RunCommand swCommands_Save, ""
            
            swModel.Extension.SetUserPreferenceToggle swUserPreferenceToggle_e.swDetailingModeSaveModelData, swUserPreferenceOption_e.swDetailingNoOptionSpecified, saveModelDataOpt
            swModel.Extension.SetUserPreferenceToggle swUserPreferenceToggle_e.swDetailingModeIncludeStandardViewsInViewPalette, swUserPreferenceOption_e.swDetailingNoOptionSpecified, includeStandardView
        Else
            Err.Raise vbError, "", "Only drawing documents are supported"
        End If
    Else
        Err.Raise vbError, "", "Open drawing document"
    End If
    
End Sub