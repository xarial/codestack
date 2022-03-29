Const BASE_NAME As String = "LinkedCustomProperties"
Const EMBED As Boolean = False
Const CLEAR_PROPERTIES As Boolean = False

Const PARAM_CSV_PATH As String = "CsvPath"
Const PARAM_CONF_SPEC_NAME As String = "ConfigurationSpecific"

Function GetParameters(app As SldWorks.SldWorks, ByRef csvFilePath As String, ByRef confSpecific As Boolean) As Boolean
    
    csvFilePath = app.GetOpenFileName("Custom Properties Template File", "", "CSV Files (*.csv)|*.csv|Text Files (*.txt)|*.txt|All Files (*.*)|*.*|", 0, "", "")

    If csvFilePath <> "" Then
        confSpecific = app.SendMsgToUser2("Link to configuration specific properties (Yes) or File Specific (No)?", swMessageBoxIcon_e.swMbQuestion, swMessageBoxBtn_e.swMbYesNo) = swMessageBoxResult_e.swMbHitYes
        GetParameters = True
    Else
        GetParameters = False
    End If
    
End Function

Sub main()

    Dim swApp As SldWorks.SldWorks
    Set swApp = Application.SldWorks
    
    Dim swModel As SldWorks.ModelDoc2
    
    Set swModel = swApp.ActiveDoc
    
    If Not swModel Is Nothing Then
                
        Dim csvFilePath As String
        Dim confSpecific As Boolean
        
        If GetParameters(swApp, csvFilePath, confSpecific) Then
                
            Dim curMacroPath As String
            curMacroPath = swApp.GetCurrentMacroPathName
            Dim vMethods(8) As String
            Dim moduleName As String
            
            GetMacroEntryPoint swApp, curMacroPath, moduleName, ""
            
            vMethods(0) = curMacroPath: vMethods(1) = moduleName: vMethods(2) = "swmRebuild"
            vMethods(3) = curMacroPath: vMethods(4) = moduleName: vMethods(5) = "swmEditDefinition"
            vMethods(6) = curMacroPath: vMethods(7) = moduleName: vMethods(8) = "swmSecurity"
            
            Dim vParamNames(1) As String
            vParamNames(0) = PARAM_CSV_PATH
            vParamNames(1) = PARAM_CONF_SPEC_NAME
    
            Dim vParamTypes(1) As Long
            vParamTypes(0) = swMacroFeatureParamType_e.swMacroFeatureParamTypeString
            vParamTypes(1) = swMacroFeatureParamType_e.swMacroFeatureParamTypeInteger
    
            Dim vParamValues(1) As String
    
            vParamValues(0) = csvFilePath
            vParamValues(1) = CLng(confSpecific)
            
            Dim opts As swMacroFeatureOptions_e
            opts = swMacroFeatureOptions_e.swMacroFeatureAlwaysAtEnd
            
            If EMBED Then
                opts = opts + swMacroFeatureOptions_e.swMacroFeatureEmbedMacroFile
            End If
            
            Dim swFeat As SldWorks.Feature
            Set swFeat = swModel.FeatureManager.InsertMacroFeature3(BASE_NAME, "", vMethods, _
                vParamNames, vParamTypes, vParamValues, Empty, Empty, Empty, _
                Empty, opts)
            
            If swFeat Is Nothing Then
                MsgBox "Failed to create linked properties feature"
            End If
            
        End If
        
    Else
        MsgBox "Please open model"
    End If
    
End Sub

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

Function GetArrayFromCsv(filePath As String) As Variant
    
    Dim fileNo As Integer

    fileNo = FreeFile
    
    Dim rows As Collection
    Set rows = New Collection
    
    Open filePath For Input As #fileNo
    
    Do While Not EOF(fileNo)
        
        Dim tableRow As String
        
        Line Input #fileNo, tableRow
            
        Dim vCells As Variant
        vCells = Split(tableRow, ",")
        rows.Add vCells
    
    Loop
    
    Close #fileNo
    
    Dim tableData() As String

    Dim rowCount As Integer
    Dim columnCount As Integer
    rowCount = rows.Count
    columnCount = UBound(rows(1)) + 1
    
    Dim rowIndex As Integer
    Dim columnIndex As Integer
    
    ReDim tableData(rowCount - 1, columnCount - 1)
    
    For rowIndex = 1 To rowCount
        Dim vRow As Variant
        vRow = rows.Item(rowIndex)
        
        For columnIndex = 1 To columnCount
            Dim cellVal As String
            cellVal = vRow(columnIndex - 1)
            
            If Left(cellVal, 2) = """""" And Right(cellVal, 2) = """""" Then
                cellVal = Mid(cellVal, 3, Len(cellVal) - 4)
            End If
            
            tableData(rowIndex - 1, columnIndex - 1) = cellVal
        Next
    Next
    
    GetArrayFromCsv = tableData
    
End Function

Function swmRebuild(varApp As Variant, varDoc As Variant, varFeat As Variant) As Variant

    Dim swApp As SldWorks.SldWorks
    Dim swModel As SldWorks.ModelDoc2
    Dim swFeat As SldWorks.Feature
    
    Set swApp = varApp
    Set swModel = varDoc
    Set swFeat = varFeat
    
    Dim swMacroFeat As SldWorks.MacroFeatureData
    Set swMacroFeat = swFeat.GetDefinition()
    
    Dim csvFilePath As String
    Dim confSpecific As Long
    
    swMacroFeat.GetStringByName PARAM_CSV_PATH, csvFilePath
    swMacroFeat.GetIntegerByName PARAM_CONF_SPEC_NAME, confSpecific
        
try_:
    On Error GoTo catch_
    
    WritePropertiesFromFile swModel, csvFilePath, IIf(CBool(confSpecific), swMacroFeat.CurrentConfiguration, Nothing)
    
    GoTo finally_
catch_:
    swmRebuild = Err.Description
finally_:
        
End Function

Sub WritePropertiesFromFile(model As SldWorks.ModelDoc2, csvFilePath As String, conf As SldWorks.Configuration)
    
    If Dir(csvFilePath) = "" Then
        Err.Raise "Linked CSV file is missing: " & csvFilePath
    End If
    
    Dim vTable As Variant
    vTable = GetArrayFromCsv(csvFilePath)
    
    Dim i As Integer
    
    Dim confName As String
    
    If conf Is Nothing Then
        confName = ""
    Else
        confName = conf.Name
    End If
    
    Dim swCustPrpMgr As SldWorks.CustomPropertyManager
    
    Set swCustPrpMgr = model.Extension.CustomPropertyManager(confName)
    
    If UBound(vTable, 2) <> 1 Then
        Err.Raise vbError, "", "There must be only 2 columns in the CSV file"
    End If
    
    If CLEAR_PROPERTIES Then
        ClearProperties swCustPrpMgr
    End If
    
    For i = 0 To UBound(vTable, 1)
                
        Dim prpName As String
        prpName = CStr(vTable(i, 0))
        
        Dim prpVal As String
        prpVal = CStr(vTable(i, 1))
        
        If swCustPrpMgr.Add3(prpName, swCustomInfoType_e.swCustomInfoText, prpVal, swCustomPropertyAddOption_e.swCustomPropertyReplaceValue) <> swCustomInfoAddResult_e.swCustomInfoAddResult_AddedOrChanged Then
            Err.Raise vbError, "", "Failed to add property '" & prpName & "'"
        End If
        
    Next
    
End Sub

Sub ClearProperties(custPrpMgr As SldWorks.CustomPropertyManager)
    
    Dim vPrpNames As Variant
    vPrpNames = custPrpMgr.GetNames
        
    If Not IsEmpty(vPrpNames) Then
        
        Dim i As Integer
        
        For i = 0 To UBound(vPrpNames)
            custPrpMgr.Delete2 CStr(vPrpNames(i))
        Next
    
    End If
    
End Sub

Function swmEditDefinition(varApp As Variant, varDoc As Variant, varFeat As Variant) As Variant

    Dim swApp As SldWorks.SldWorks
    Set swApp = varApp

    Dim csvFilePath As String
    Dim confSpecific As Boolean
    
    If GetParameters(swApp, csvFilePath, confSpecific) Then
        
        Dim swModel As SldWorks.ModelDoc2
        Dim swFeat As SldWorks.Feature
        
        Set swModel = varDoc
        Set swFeat = varFeat
        
        Dim swMacroFeat As SldWorks.MacroFeatureData
        Set swMacroFeat = swFeat.GetDefinition()
        
        swMacroFeat.AccessSelections swModel, Nothing
        swMacroFeat.SetStringByName PARAM_CSV_PATH, csvFilePath
        swMacroFeat.SetIntegerByName PARAM_CONF_SPEC_NAME, CLng(confSpecific)

        swFeat.ModifyDefinition swMacroFeat, swModel, Nothing
        
    End If
    
    swmEditDefinition = True
    
End Function

Function swmSecurity(varApp As Variant, varDoc As Variant, varFeat As Variant) As Variant
    swmSecurity = SwConst.swMacroFeatureSecurityOptions_e.swMacroFeatureSecurityByDefault
End Function