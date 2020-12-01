Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swModel As SldWorks.ModelDoc2
    
    Set swModel = swApp.ActiveDoc
    
    Debug.Print GetPropertyValue(swModel, "Part Number")
    Debug.Print GetPropertyValue(swModel, "Revision")
    
End Sub

Function GetPropertyValue(model As SldWorks.ModelDoc2, prpName As String) As String
    
    Dim prpVal As String
    Dim swCustPrpMgr As SldWorks.CustomPropertyManager
    
    If TypeOf model Is SldWorks.PartDoc Or TypeOf model Is SldWorks.AssemblyDoc Then
        Set swCustPrpMgr = model.ConfigurationManager.ActiveConfiguration.CustomPropertyManager
        swCustPrpMgr.Get4 prpName, True, "", prpVal
    End If
    
    If prpVal = "" Then
        Set swCustPrpMgr = model.Extension.CustomPropertyManager("")
        swCustPrpMgr.Get4 prpName, True, "", prpVal
    End If
    
    GetPropertyValue = prpVal
    
End Function