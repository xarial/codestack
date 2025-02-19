Type DimensionInfo
    Name As String
    title As String
    Value As Double
End Type

Public Const MARGIN As Integer = 10
Public Const MAX_FORM_HEIGHT = 200
Public Const TEXT_BOX_WIDTH As Integer = 50
Public Const BASE_NAME As String = "Configurator"

Const EMBED_MACRO_FEATURE As Boolean = False

Sub main()

    Dim swApp As SldWorks.SldWorks
    Set swApp = Application.SldWorks

try_:

    On Error GoTo catch_
    
    Dim swModel As SldWorks.ModelDoc2
    
    Set swModel = swApp.ActiveDoc
    
    If Not swModel Is Nothing Then
        
        If Not TypeOf swModel Is PartDoc And Not TypeOf swModel Is AssemblyDoc Then
            Err.Raise vbError, "", "Only part and assembly documents are supported"
        End If
        
        Dim vParamNames As Variant
        Dim vParamTypes As Variant
        Dim vParamValues As Variant
        
        If Not CollectParameters(swModel, vParamNames, vParamTypes, vParamValues) Then
            Err.Raise vbError, "", "Please select dimensions to configure"
        End If
        
        Dim curMacroPath As String
        curMacroPath = swApp.GetCurrentMacroPathName
        
        Dim vMethods(8) As String
        Dim moduleName As String
        
        GetMacroEntryPoint swApp, curMacroPath, moduleName, ""
        
        vMethods(0) = curMacroPath: vMethods(1) = moduleName: vMethods(2) = "swmRebuild"
        vMethods(3) = curMacroPath: vMethods(4) = moduleName: vMethods(5) = "swmEditDefinition"
        vMethods(6) = curMacroPath: vMethods(7) = moduleName: vMethods(8) = "swmSecurity"
        
        Dim opts As swMacroFeatureOptions_e
        
        If EMBED_MACRO_FEATURE Then
            opts = swMacroFeatureOptions_e.swMacroFeatureEmbedMacroFile
        Else
            opts = swMacroFeatureOptions_e.swMacroFeatureByDefault
        End If
        
        Dim swFeat As SldWorks.Feature
        Set swFeat = swModel.FeatureManager.InsertMacroFeature3(BASE_NAME, "", vMethods, _
            vParamNames, vParamTypes, vParamValues, Empty, Empty, Empty, _
            Empty, opts)
        
        If swFeat Is Nothing Then
            Err.Raise vbError, "", "Failed to create box feature"
        End If
        
    Else
        Err.Raise "Please open model"
    End If
    
    GoTo finally_
    
catch_:
    MsgBox Err.Description, vbCritical, "Configurator"
finally_:
    
End Sub

Function CollectParameters(model As SldWorks.ModelDoc2, ByRef vParamNames As Variant, ByRef vParamTypes As Variant, ByRef vParamValues As Variant) As Boolean

    Dim paramNames() As String
    Dim paramTypes() As Long
    Dim paramValues() As String

    Dim swSelMgr As SldWorks.SelectionMgr
    Set swSelMgr = model.SelectionManager

    Dim i As Integer
    
    For i = 1 To swSelMgr.GetSelectedObjectCount2(-1)
        If swSelMgr.GetSelectedObjectType3(i, -1) = swSelectType_e.swSelDIMENSIONS Then
            
            Dim swDispDim As SldWorks.DisplayDimension
            Set swDispDim = swSelMgr.GetSelectedObject6(i, -1)
            
            Dim swComp As SldWorks.Component2
            Set swComp = swSelMgr.GetSelectedObjectsComponent3(i, -1)
                        
            If (Not paramNames) = -1 Then
                ReDim paramNames(0)
                ReDim paramTypes(0)
                ReDim paramValues(0)
            Else
                ReDim Preserve paramNames(UBound(paramNames) + 1)
                ReDim Preserve paramTypes(UBound(paramTypes) + 1)
                ReDim Preserve paramValues(UBound(paramValues) + 1)
            End If
            
            Dim paramName As String
            paramName = ""
            
            If Not swComp Is Nothing Then
                
                paramName = swComp.Name2
                
                Dim swAssy As SldWorks.AssemblyDoc
                Set swAssy = model
                
                Dim swEditTargetComp As SldWorks.Component2
                Set swEditTargetComp = swAssy.GetEditTargetComponent
                
                If Not swEditTargetComp Is Nothing Then
                    If Not swEditTargetComp.GetModelDoc2() Is swAssy Then
                        If Left(paramName, Len(swEditTargetComp.Name2)) <> swEditTargetComp.Name2 Then
                            Err.Raise vbError, "", "Dimension must belong to the current edit target"
                        End If
                        If LCase(paramName) = LCase(swEditTargetComp.Name2) Then
                            paramName = ""
                        Else
                            paramName = Right(paramName, Len(paramName) - Len(swEditTargetComp.Name2) - 1)
                        End If
                    End If
                End If
                
            End If
            
            paramName = paramName & IIf(paramName <> "", "/", "") & swDispDim.GetNameForSelection
            
            paramNames(UBound(paramNames)) = paramName
            paramValues(UBound(paramValues)) = InputBox("Specify the name for " & paramName, "Configurator", paramName)
            paramTypes(UBound(paramTypes)) = swMacroFeatureParamType_e.swMacroFeatureParamTypeString
            
        End If
    Next
    
    vParamNames = paramNames
    vParamTypes = paramTypes
    vParamValues = paramValues
    
    CollectParameters = (Not paramNames) <> -1
    
End Function

Sub GetMacroEntryPoint(app As SldWorks.SldWorks, macroPath As String, ByRef moduleName As String, ByRef procName As String)
        
    Dim vMethods As Variant
    vMethods = app.GetMacroMethods(macroPath, swMacroMethods_e.swMethodsWithoutArguments)
    
    Dim i As Integer
    
    If Not IsEmpty(vMethods) Then
    
        For i = 0 To UBound(vMethods)
            Dim vData As Variant
            vData = Split(vMethods(i), ".")
            
            If i = 0 Or LCase(vData(1)) = "main" Then
                moduleName = vData(0)
                procName = vData(1)
            End If
        Next
        
    End If
    
End Sub

Function swmRebuild(varApp As Variant, varDoc As Variant, varFeat As Variant) As Variant
    swmRebuild = True
End Function

Function swmSecurity(varApp As Variant, varDoc As Variant, varFeat As Variant) As Variant
    swmSecurity = SwConst.swMacroFeatureSecurityOptions_e.swMacroFeatureSecurityByDefault
End Function

Function swmEditDefinition(varApp As Variant, varDoc As Variant, varFeat As Variant) As Variant
    
try_:

    On Error GoTo catch_

    Dim swFeat As SldWorks.Feature
    Set swFeat = varFeat
    
    Dim title As String
    title = "Edit " & swFeat.Name
    
    Dim swMacroFeat As SldWorks.MacroFeatureData
    Set swMacroFeat = swFeat.GetDefinition
        
    Dim vParamNames As Variant
    Dim vParamValues As Variant
    
    swMacroFeat.GetParameters vParamNames, Empty, vParamValues
        
    Dim swActiveModel As SldWorks.ModelDoc2
    
    Set swActiveModel = varDoc
    
    Dim confName As String
    confName = swMacroFeat.CurrentConfiguration.Name
    
    Dim dimsInfo() As DimensionInfo
    dimsInfo = LoadDimensionValues(swActiveModel, confName, vParamNames, vParamValues)
    
    ConfiguratorForm.Caption = title
    
    ConfiguratorForm.EditDimensions dimsInfo, swActiveModel, confName
        
    swmEditDefinition = True
        
    GoTo finally_
    
catch_:
    swmEditDefinition = False
    MsgBox Err.Description, vbCritical, title
finally_:

End Function

Public Sub TrySetDimensions(dimsInfo() As DimensionInfo, model As SldWorks.ModelDoc2, targConfName As String, createConf As Boolean)
    
try_:

    On Error GoTo catch_
    
    Dim swTargModel As SldWorks.ModelDoc2
    Dim swTargComp As SldWorks.Component2
        
    If model.GetType() = swDocumentTypes_e.swDocASSEMBLY Then
        Dim swAssy As SldWorks.AssemblyDoc
        Set swAssy = model
        Set swTargModel = swAssy.GetEditTarget
        Set swTargComp = swAssy.GetEditTargetComponent
    Else
        Set swTargModel = model
    End If
    
    If createConf Then
        Dim swConf As SldWorks.Configuration
                
        If targConfName = "" Then
            Err.Raise vbError, "", "Specify configuration name"
        End If
        
        Set swConf = swTargModel.ConfigurationManager.AddConfiguration2(targConfName, "", "", swConfigurationOptions2_e.swConfigOption_DontActivate, "", "", False)
        If swConf Is Nothing Then
            Err.Raise vbError, "", "Failed to add new configuration"
        End If
    End If
    
    Dim i As Integer
        
    For i = 0 To UBound(dimsInfo)
        
        Dim dimInfo As DimensionInfo
        dimInfo = dimsInfo(i)
        
        Dim swDim As SldWorks.Dimension
        
        Dim dimName As String
        dimName = dimInfo.Name
        
        Set swDim = GetDimension(swTargModel, dimName)
        
        If Not swDim Is Nothing Then
            Dim dimVal As Double
            dimVal = dimInfo.Value
            
            Dim confNames(0) As String
            confNames(0) = targConfName
            swDim.SetValue3 dimVal, swInConfigurationOpts_e.swSpecifyConfiguration, confNames
        Else
            Err.Raise vbError, "", dimName & " does not exist"
        End If
    Next
    
    If createConf And Not swTargComp Is Nothing Then
        
        swTargComp.ReferencedConfiguration = targConfName
        
    End If
    
    GoTo finally_
    
catch_:
    MsgBox Err.Description, vbCritical, "Configurator"
finally_:
    
End Sub

Function GetDimension(model As SldWorks.ModelDoc2, dimName As String) As SldWorks.Dimension
    
    Dim dimParts As Variant
    dimParts = Split(dimName, "/")
    
    Dim i As Integer
    
    Dim swTargetModel As SldWorks.ModelDoc2
    Set swTargetModel = model
    
    Dim swCurComp As SldWorks.Component2
    
    For i = 0 To UBound(dimParts) - 1
        Dim swAssy As SldWorks.AssemblyDoc
        Set swAssy = swTargetModel
        Set swCurComp = swAssy.GetComponentByName(dimParts(i))
        Set swTargetModel = swCurComp.GetModelDoc2()
    Next
    
    Set GetDimension = swTargetModel.Parameter(dimParts(UBound(dimParts)))
    
End Function

Private Function LoadDimensionValues(model As SldWorks.ModelDoc2, confName As String, vParamNames As Variant, vParamValues As Variant) As DimensionInfo()

    Dim swTargModel As SldWorks.ModelDoc2

    If model.GetType() = swDocumentTypes_e.swDocASSEMBLY Then
        Dim swAssy As SldWorks.AssemblyDoc
        Set swAssy = model
        Set swTargModel = swAssy.GetEditTarget
    Else
        Set swTargModel = model
    End If

    Dim dimsInfo() As DimensionInfo
    ReDim dimsInfo(UBound(vParamNames))

    Dim i As Integer

    For i = 0 To UBound(vParamNames)

        Dim swDim As SldWorks.Dimension

        Dim dimName As String
        dimName = CStr(vParamNames(i))
        
        dimsInfo(i).Name = dimName
        dimsInfo(i).title = vParamValues(i)

        Set swDim = GetDimension(swTargModel, dimName)

        If Not swDim Is Nothing Then
            Dim dimVal As Double
            Dim confNames(0) As String
            confNames(0) = confName
            dimVal = swDim.GetValue3(swInConfigurationOpts_e.swSpecifyConfiguration, confNames)(0)
            dimsInfo(i).Value = dimVal
        Else
            Err.Raise vbError, "", dimName & " does not exist"
        End If
    Next
    
    LoadDimensionValues = dimsInfo

End Function