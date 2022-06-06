#Const ARGS = True 'True to use arguments from Toolbar+ or Batch+ instead of the constant

Const CLEAR_PROPERTIES As Boolean = False

Sub main()
    
    Dim swApp As SldWorks.SldWorks
    Set swApp = Application.SldWorks
    
    Dim swModel As SldWorks.ModelDoc2
    
try_:
    On Error GoTo catch_
                
    Dim csvFilePath As String
    Dim confSpecific As Boolean
    
    If GetParameters(swApp, swModel, csvFilePath, confSpecific) Then
    
        If Not swModel Is Nothing Then
            WritePropertiesFromFile swModel, csvFilePath, IIf(CBool(confSpecific), swModel.ConfigurationManager.ActiveConfiguration, Nothing)
        Else
            Err.Raise vbError, "", "Please open model"
        End If
        
    End If
            
    GoTo finally_
catch_:
    swmRebuild = Err.Description
finally_:
    
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