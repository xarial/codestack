Const SUPPRESS_UPDATES As Boolean = True

Const SRC_PART As String = "C:\Sample.sldprt"

Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swModel As SldWorks.ModelDoc2
    
    Set swModel = swApp.ActiveDoc
    
    If TypeOf swModel Is SldWorks.PartDoc Then
        
        On Error GoTo End_

        If SUPPRESS_UPDATES Then
            SuppressUpdates swModel, True
        End If
        
        Dim activeConfName As String
        activeConfName = swModel.ConfigurationManager.ActiveConfiguration.Name
        
        Dim vBodies As Variant
        vBodies = GetBodies(SRC_PART)
        
        swModel.ConfigurationManager.AddConfiguration2 activeConfName & "_Merged", "", "", swConfigurationOptions2_e.swConfigOption_LinkToParent, activeConfName, "", True
        
        Dim i As Integer
        
        For i = 0 To UBound(vBodies)
            Dim swBody As SldWorks.Body2
            Set swBody = vBodies(i)
            Dim swFeat As SldWorks.Feature
            Set swFeat = swModel.CreateFeatureFromBody3(swBody, False, swCreateFeatureBodyOpts_e.swCreateFeatureBodySimplify)
            swFeat.SetSuppression2 swFeatureSuppressionAction_e.swUnSuppressFeature, swInConfigurationOpts_e.swThisConfiguration, Empty
        Next
        
        swModel.ShowConfiguration2 activeConfName

End_: 'restore the flag otherwise all files will be opened invisible
    
        If SUPPRESS_UPDATES Then
            SuppressUpdates swModel, False
        End If
        
    Else
        MsgBox "Please open part document"
    End If
    
End Sub

Sub SuppressUpdates(model As SldWorks.ModelDoc2, suppress As Boolean)
    
    Dim enable As Boolean
    enable = Not suppress
    
    Dim swView As SldWorks.ModelView
    Set swView = model.ActiveView
    
    swView.EnableGraphicsUpdate = enable
    
    model.FeatureManager.EnableFeatureTree = enable
    model.FeatureManager.EnableFeatureTreeWindow = enable
        
    swApp.DocumentVisible enable, swDocumentTypes_e.swDocPART
    swApp.DocumentVisible enable, swDocumentTypes_e.swDocASSEMBLY
    swApp.DocumentVisible enable, swDocumentTypes_e.swDocDRAWING
    
End Sub

Function GetBodies(path As String) As Variant
    
    Dim swPart As SldWorks.PartDoc
    Set swPart = swApp.OpenDoc6(path, swDocumentTypes_e.swDocPART, _
        swOpenDocOptions_e.swOpenDocOptions_Silent + swOpenDocOptions_e.swOpenDocOptions_ReadOnly, "", 0, 0)
    
    Dim vBodies As Variant
    vBodies = swPart.GetBodies2(swBodyType_e.swSolidBody, True)
    
    If Not IsEmpty(vBodies) Then
        Dim i As Integer
        For i = 0 To UBound(vBodies)
            Dim swBody As SldWorks.Body2
            Set swBody = vBodies(i)
            Set vBodies(i) = swBody.Copy
        Next
    End If
    
    swApp.CloseDoc swPart.GetTitle()
    
    GetBodies = vBodies
    
End Function