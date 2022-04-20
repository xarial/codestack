Const UPDATE_ALL_COMPS As Boolean = True
Const REBUILD_ALL_CONFIGS As Boolean = True

Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swModel As SldWorks.ModelDoc2
    
    Set swModel = swApp.ActiveDoc
    
    swModel.Extension.UpgradeLegacyCustomProperties UPDATE_ALL_COMPS
    
    If REBUILD_ALL_CONFIGS Then
        swModel.Extension.ForceRebuildAll
    End If
    
End Sub