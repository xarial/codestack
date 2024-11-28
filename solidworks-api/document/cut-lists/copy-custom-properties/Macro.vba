Const CONF_SPEC_PRP As Boolean = True
Const COPY_RES_VAL As Boolean = True

Const ALL_CONFS As Boolean = False
Const PROCESS_TOP_LEVEL_CONFIGS As Boolean = False
Const PROCESS_CHILDREN_CONFIGS As Boolean = True

Dim SRC_PROPERTIES As Variant
Dim TARG_PROPERTIES As Variant

Dim swApp As SldWorks.SldWorks

Sub Init(Optional dummy As Variant = Empty)    
    
    SRC_PROPERTIES = Array("Bounding Box Length", "Bounding Box Width", "Sheet Metal Thickness") 'list of custom properties to copy or Empty to copy all
    TARG_PROPERTIES = Array("Length", "Width", "Thickness") 'list of target custom property namesor Empty to use original name

End Sub

Sub main()
    
try_:
    
    On Error GoTo catch_
    
    Init
    
    Set swApp = Application.SldWorks
    
    Dim swModel As SldWorks.ModelDoc2
    Set swModel = swApp.ActiveDoc
    
    Dim activeConfName As String
    activeConfName = swModel.ConfigurationManager.ActiveConfiguration.Name
    
    Dim vConfNames As Variant
    vConfNames = GetConfigurations(swModel)
    
    Dim i As Integer
    
    For i = 0 To UBound(vConfNames)
        
        swModel.ShowConfiguration2 CStr(vConfNames(i))
        
        Dim swCutListPrpMgr As SldWorks.CustomPropertyManager
        Set swCutListPrpMgr = GetCutListPropertyManager(swModel)
        
        If Not swCutListPrpMgr Is Nothing Then
            
            Dim swTargetPrpMgr As SldWorks.CustomPropertyManager
            
            If CONF_SPEC_PRP Then
                Set swTargetPrpMgr = swModel.ConfigurationManager.ActiveConfiguration.CustomPropertyManager
            Else
                Set swTargetPrpMgr = swModel.Extension.CustomPropertyManager("")
            End If
            
            CopyProperties swCutListPrpMgr, swTargetPrpMgr, SRC_PROPERTIES, TARG_PROPERTIES
            
        Else
            Err.Raise vbError, "", "Cut-list is not found"
        End If
    
    Next
    
    GoTo finally_
    
catch_:
    swApp.SendMsgToUser2 Err.Description, swMessageBoxIcon_e.swMbStop, swMessageBoxBtn_e.swMbOk
finally_:

    If activeConfName <> "" Then
        swModel.ShowConfiguration2 activeConfName
    End If

End Sub

Function GetCutListPropertyManager(model As SldWorks.ModelDoc2) As SldWorks.CustomPropertyManager

    Dim swFeat As SldWorks.Feature
    
    Set swFeat = model.FirstFeature
    
    While Not swFeat Is Nothing
        
        If swFeat.GetTypeName2() = "CutListFolder" Then
 
            Dim swBodyFolder As SldWorks.BodyFolder
            Set swBodyFolder = swFeat.GetSpecificFeature2
 
            Dim bodyCount As Long
            bodyCount = swBodyFolder.GetBodyCount
 
            If bodyCount <> 0 Then
                Set GetCutListPropertyManager = swFeat.CustomPropertyManager
                Exit Function
            End If
        End If

        Set swFeat = swFeat.GetNextFeature
        
    Wend
    
End Function

Sub CopyProperties(srcPrpMgr As SldWorks.CustomPropertyManager, targPrpMgr As SldWorks.CustomPropertyManager, vSrcPrpNames As Variant, vTargPrpNames As Variant)

    If IsEmpty(vSrcPrpNames) Then
        vSrcPrpNames = srcPrpMgr.GetNames()
        vTargPrpNames = vSrcPrpNames
    End If
    
    If IsEmpty(vTargPrpNames) Then
        vTargPrpNames = vSrcPrpNames
    End If
    
    If Not IsEmpty(vSrcPrpNames) Then
        
        If UBound(vSrcPrpNames) = UBound(vTargPrpNames) Then
        
            For i = 0 To UBound(vSrcPrpNames)
                            
                Dim srcPrpName As String
                
                srcPrpName = vSrcPrpNames(i)
    
                Dim prpVal As String
                Dim prpResVal As String
                            
                srcPrpMgr.Get5 srcPrpName, False, prpVal, prpResVal, False
                
                Dim targVal As String
                targVal = IIf(COPY_RES_VAL, prpResVal, prpVal)
                
                Dim targPrpName As String
                
                targPrpName = vTargPrpNames(i)
                
                targPrpMgr.Add2 targPrpName, swCustomInfoType_e.swCustomInfoText, targVal
                targPrpMgr.Set targPrpName, targVal
                
            Next
        Else
            Err.Raise vbError, "", "Target proeprties name do not match source"
        End If
        
    Else
        Err.Raise vbError, "", "No properties to copy"
    End If
    
End Sub

Function GetConfigurations(model As SldWorks.ModelDoc2) As Variant
    
    Dim confNames() As String
    
    If ALL_CONFS And CONF_SPEC_PRP Then
    
        Dim vConfNames As Variant
        vConfNames = model.GetConfigurationNames
        
        Dim i As Integer
        
        For i = 0 To UBound(vConfNames)
            
            Dim confName As String
            confName = CStr(vConfNames(i))
            
            Dim swConf As SldWorks.Configuration
            Set swConf = model.GetConfigurationByName(confName)
            
            If swConf.Type = swConfigurationType_e.swConfiguration_Standard Then
                    
                If (PROCESS_TOP_LEVEL_CONFIGS And swConf.GetParent() Is Nothing) Or (PROCESS_CHILDREN_CONFIGS And Not swConf.GetParent() Is Nothing) Then
                    If (Not confNames) = -1 Then
                        ReDim confNames(0)
                    Else
                        ReDim Preserve confNames(UBound(confNames) + 1)
                    End If
                
                    confNames(UBound(confNames)) = confName
                
                End If
            
            End If
            
        Next
    
    Else
        ReDim confNames(0)
        confNames(0) = model.ConfigurationManager.ActiveConfiguration.Name
    End If
    
    GetConfigurations = confNames
    
End Function
