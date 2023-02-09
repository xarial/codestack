Type RefCompModel
    RefModel As SldWorks.ModelDoc2
    RefConf As String
End Type

#Const ARGS = True 'True to use arguments from Toolbar+ or Batch+ instead of the constant

Const CLEAR_PROPERTIES As Boolean = False
Const ALL_COMPONENTS As Boolean = False

Sub main()
    
    Dim swApp As SldWorks.SldWorks
    Set swApp = Application.SldWorks
    
    Dim swModel As SldWorks.ModelDoc2
                
    Dim csvFilePath As String
    Dim confSpecific As Boolean
    
    If GetParameters(swApp, swModel, csvFilePath, confSpecific) Then
    
        If Not swModel Is Nothing Then
            
            Dim vTable As Variant
            vTable = GetArrayFromCsv(csvFilePath)
            
            Dim swRefConf As SldWorks.Configuration
            Set swRefConf = swModel.ConfigurationManager.ActiveConfiguration
            
            WritePropertiesFromTable swModel, vTable, IIf(CBool(confSpecific), swRefConf.Name, ""), CLEAR_PROPERTIES
        
            If ALL_COMPONENTS Then
            
                Dim refCompModels() As RefCompModel
                refCompModels = CollectUniqueComponents(swRefConf, confSpecific)
                
                If (Not refCompModels) <> -1 Then
                    
                    Dim i As Integer
                    
                    For i = 0 To UBound(refCompModels)
                        WritePropertiesFromTable refCompModels(i).RefModel, vTable, refCompModels(i).RefConf, CBool(clearPrps)
                    Next
                    
                End If
            
            End If
        
            'WritePropertiesFromFile swModel, csvFilePath, IIf(CBool(confSpecific), swModel.ConfigurationManager.ActiveConfiguration, Nothing)
        Else
            Err.Raise vbError, "", "Please open model"
        End If
        
    End If
            
End Sub

Function GetParameters(app As SldWorks.SldWorks, ByRef model As SldWorks.ModelDoc2, ByRef csvFilePath As String, ByRef confSpecific As Boolean) As Boolean
    
Dim confSpecArgsParsed As Boolean

#If ARGS Then

try_:
    On Error GoTo catch_
    
    Dim macroRunner As Object
    Set macroRunner = CreateObject("CadPlus.MacroRunner.Sw")
    
    Dim param As Object
    Set param = macroRunner.PopParameter(app)
        
    Dim vArgs As Variant
    vArgs = param.Get("Args")
        
    Set model = param.Get("Model")
    
    If Not IsEmpty(vArgs) Then
        csvFilePath = CStr(vArgs(0))
    End If
    
    If UBound(vArgs) > 0 Then
        confSpecific = CBool(vArgs(1))
        confSpecArgsParsed = True
    End If
    
    GoTo finally_
    
catch_:
finally_:

#End If

    If Trim(csvFilePath) = "" Then
        csvFilePath = app.GetOpenFileName("Custom Properties Template File", "", "CSV Files (*.csv)|*.csv|Text Files (*.txt)|*.txt|All Files (*.*)|*.*|", 0, "", "")
    End If
    
    If model Is Nothing Then
        Set model = app.ActiveDoc
    End If
    
    If csvFilePath <> "" Then
        If Not confSpecArgsParsed Then
            confSpecific = app.SendMsgToUser2("Link to configuration specific properties (Yes) or File Specific (No)?", swMessageBoxIcon_e.swMbQuestion, swMessageBoxBtn_e.swMbYesNo) = swMessageBoxResult_e.swMbHitYes
        End If
        GetParameters = True
    Else
        GetParameters = False
    End If
    
End Function

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

Sub WritePropertiesFromTable(model As SldWorks.ModelDoc2, table As Variant, confName As String, clearPrps As Boolean)
    
    Dim i As Integer
    
    Dim swCustPrpMgr As SldWorks.CustomPropertyManager
    
    Set swCustPrpMgr = model.Extension.CustomPropertyManager(confName)
    
    If clearPrps Then
        ClearProperties swCustPrpMgr
    End If
    
    For i = 0 To UBound(table, 1)
                
        Dim prpName As String
        prpName = CStr(table(i, 0))
        
        Dim prpVal As String
        prpVal = CStr(table(i, 1))
        
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

Function CollectUniqueComponents(assmConf As SldWorks.Configuration, confSpecific As Boolean) As RefCompModel()
    
    Dim swRootComp As SldWorks.Component2
    Set swRootComp = assmConf.GetRootComponent3(False)
    
    Dim refCompModels() As RefCompModel
    
    ProcessComponents swRootComp.GetChildren(), confSpecific, refCompModels
    
    CollectUniqueComponents = refCompModels
    
End Function

Sub ProcessComponents(vComps As Variant, confSpecific As Boolean, refCompModels() As RefCompModel)
    
    If Not IsEmpty(vComps) Then
    
        Dim i As Integer
        
        For i = 0 To UBound(vComps)
            
            Dim swComp As SldWorks.Component2
            Set swComp = vComps(i)
            
            Dim swRefModel As SldWorks.ModelDoc2
            Set swRefModel = swComp.GetModelDoc2
            
            If Not swRefModel Is Nothing Then
            
                Dim refConfName As String
                
                refConfName = IIf(confSpecific, swComp.ReferencedConfiguration, "")
                
                If Not Contains(refCompModels, swRefModel, refConfName) Then
                
                    If (Not refCompModels) = -1 Then
                        ReDim refCompModels(0)
                    Else
                        ReDim Preserve refCompModels(UBound(refCompModels) + 1)
                    End If
                    
                    Set refCompModels(UBound(refCompModels)).RefModel = swRefModel
                    refCompModels(UBound(refCompModels)).RefConf = refConfName
                    
                End If
                
                ProcessComponents swComp.GetChildren(), confSpecific, refCompModels
                
            End If
            
        Next
    
    End If
    
End Sub

Function Contains(refCompModels() As RefCompModel, model As SldWorks.ModelDoc2, conf As String) As Boolean
    
    Contains = False
    
    If (Not refCompModels) <> -1 Then
        
        Dim i As Integer
        
        For i = 0 To UBound(refCompModels)
                
            If refCompModels(i).RefModel Is model And LCase(refCompModels(i).RefConf) = LCase(conf) Then
                Contains = True
                Exit Function
            End If
                
        Next
        
    End If
    
End Function