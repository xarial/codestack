Option Compare Text

Const DISPLAY_STATE_NAME As String = "My Display State"

Dim swApp As SldWorks.SldWorks

Sub main()
        
    Set swApp = Application.SldWorks
        
    Dim swModel As SldWorks.ModelDoc2
    
    Set swModel = swApp.ActiveDoc
    
    Dim swConf As SldWorks.Configuration
    
    Set swConf = swModel.ConfigurationManager.ActiveConfiguration
    
    If Not Contains(swConf.GetDisplayStates(), DISPLAY_STATE_NAME) Then
        If False = swConf.CreateDisplayState(DISPLAY_STATE_NAME) Then
            Err.Raise vbError, "", "Failed to create display state"
        End If
    End If
    
    If False <> swConf.ApplyDisplayState(DISPLAY_STATE_NAME) Then
        
        Dim swSelMgr As SldWorks.SelectionMgr
        Set swSelMgr = swModel.SelectionManager
        
        Dim swFaces(0) As SldWorks.Face2
        Set swFaces(0) = swSelMgr.GetSelectedObject6(1, -1)
        
        Dim swDispStateSetts As SldWorks.DisplayStateSetting
        
        Set swDispStateSetts = swModel.Extension.GetDisplayStateSetting(swDisplayStateOpts_e.swSpecifyDisplayState)
        swDispStateSetts.Entities = swFaces
        swDispStateSetts.Option = swDisplayStateOpts_e.swSpecifyDisplayState
        
        Dim dispStateNames(0) As String
        dispStateNames(0) = DISPLAY_STATE_NAME
        swDispStateSetts.Names = dispStateNames
        
        swDispStateSetts.PartLevel = False
        
        Dim vAppearances As Variant
        
        vAppearances = swModel.Extension.DisplayStateSpecMaterialPropertyValues(swDispStateSetts)
        
        Dim swAppearanceSetts(0) As SldWorks.AppearanceSetting
        
        Set swAppearanceSetts(0) = vAppearances(0)
        
        swAppearanceSetts(0).Color = RGB(0, 255, 0)
        swAppearanceSetts(0).Diffuse = 1#
        swAppearanceSetts(0).Specular = 0.5
        swAppearanceSetts(0).SpecularColor = RGB(0, 255, 0)
        swAppearanceSetts(0).Luminous = 0.4
        swAppearanceSetts(0).Transparent = 0#
    
        swModel.Extension.DisplayStateSpecMaterialPropertyValues(swDispStateSetts) = swAppearanceSetts
        
        swModel.EditRebuild3
        
    End If
    
End Sub

Function Contains(arr As Variant, item As Variant) As Boolean
    
    If Not IsEmpty(arr) Then
    
        Dim i As Integer
        
        For i = 0 To UBound(arr)
                    
            If TypeOf item Is Object  Then
                Contains = arr(i) Is item
            Else
                Contains = arr(i) = item
            End If
        
            If Contains Then
                Exit Function
            End If
        Next
        
    End If
    
    Contains = False
    
End Function