#Const ARGS = False 'True to use arguments from Toolbar+ or Batch+ instead of the constant

Const OUT_FILE_PATH_TEMPLATE As String = "<_FileName_>-<_TableName_>.csv" 'ouput file path template
Const INCLUDE_HEADER As Boolean = True
Const TABLE_TYPE As Integer = -1  '-1 to use selected table or table type as defined in swTableAnnotationType_e
Const ALL_SHEETS As Boolean = True 'True to export from all sheets (if TABLE_TYPE is not -1), False to export from active sheet only

Const MERGE As Boolean = False

Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks

try_:
    On Error GoTo catch_
    
    Dim tableType As swTableAnnotationType_e
    Dim outFilePathTemplate As String
    
    #If ARGS Then
                
        Dim macroRunner As Object
        Set macroRunner = CreateObject("CadPlus.MacroRunner.Sw")
        
        Dim param As Object
        Set param = macroRunner.PopParameter(swApp)
        
        Dim vArgs As Variant
        vArgs = param.Get("Args")
        
        Dim operation As String
        operation = CStr(vArgs(0))
        
        Select Case LCase(operation)
            Case "-bom"
                tableType = swTableAnnotation_BillOfMaterials
            Case "-general"
                tableType = swTableAnnotation_General
            Case "-revision"
                tableType = swTableAnnotation_RevisionBlock
            Case "-cutlist"
                tableType = swTableAnnotation_WeldmentCutList
            Case Else
                Err.Raise vbError, "", "Invalid argument. Valid arguments -bom -general -revision -cutlist"
        End Select
        
        If UBound(vArgs) = 1 Then
            outFilePathTemplate = CStr(vArgs(1))
        Else
            outFilePathTemplate = OUT_FILE_PATH_TEMPLATE
        End If
    #Else
        tableType = TABLE_TYPE
        outFilePathTemplate = OUT_FILE_PATH_TEMPLATE
    #End If
    
    Dim swModel As SldWorks.ModelDoc2
    
    Set swModel = swApp.ActiveDoc
    
    If Not swModel Is Nothing Then
        
        Dim vTables As Variant
        
        If tableType = -1 Then
            vTables = GetSelectedTables(swModel)
        Else
            
            If swModel.GetType() <> swDocumentTypes_e.swDocDRAWING Then
                Err.Raise vbError, "", "Only drawing document is supported"
            End If
            
            Dim swDraw As SldWorks.DrawingDoc
            Set swDraw = swModel
            
            Dim sheetName As String
            
            If ALL_SHEETS Then
                sheetName = ""
            Else
                sheetName = swDraw.GetCurrentSheet().GetName
            End If
            
            vTables = FindTables(swDraw, tableType, sheetName)
            
        End If
        
        If Not IsEmpty(vTables) Then
            
            Dim i As Integer
            
            Dim outFilePath As String
            
            For i = 0 To UBound(vTables)
                    
                Dim swTableAnn As SldWorks.TableAnnotation
                Set swTableAnn = vTables(i)
                
                If i = 0 Or Not MERGE Then
                    outFilePath = GetExportFilePath(outFilePathTemplate, swModel, swTableAnn)
                End If
                
                Dim vTableData As Variant
                
                Dim includeHeader As Boolean
                includeHeader = INCLUDE_HEADER And (Not MERGE Or i = 0)
                
                vTableData = GetTableData(swTableAnn, includeHeader)
                
                Dim append As Boolean
                append = IIf(MERGE, i > 0, False)
                
                If MERGE And i > 0 Then
                    Dim separatorRow() As String
                    ReDim separatorRow(0, UBound(vTableData, 2))
                    WriteCsvFile outFilePath, separatorRow, True
                End If
                
                WriteCsvFile outFilePath, vTableData, append
            
            Next
            
            GoTo finally_
            
        Else
            Err.Raise vbError, "", "Tables are not found"
        End If
        
    Else
        Err.Raise vbError, "", "Document is not open"
    End If

catch_:
    swApp.SendMsgToUser2 Err.Description, swMessageBoxIcon_e.swMbStop, swMessageBoxBtn_e.swMbOk
finally_:

End Sub

Function GetExportFilePath(pathTemplate As String, model As SldWorks.ModelDoc2, tableAnn As SldWorks.TableAnnotation) As String
    
    Const FILE_NAME_TOKEN As String = "<_FileName_>"
    Const TABLE_NAME_TOKEN As String = "<_TableName_>"
    Const SHEET_NAME_TOKEN As String = "<_SheetName_>"
        
    Dim path As String
    
    path = pathTemplate
    
    If InStr(path, FILE_NAME_TOKEN) > 0 Then
        path = Replace(pathTemplate, FILE_NAME_TOKEN, GetFileNameWithoutExtension(model.GetPathName()))
    End If
    
    If InStr(path, SHEET_NAME_TOKEN) > 0 Then
        Dim swSheet As SldWorks.Sheet
        Set swSheet = GetSheetFromTableAnnotation(model, tableAnn)
        path = Replace(path, SHEET_NAME_TOKEN, swSheet.GetName())
    End If
    
    If InStr(path, TABLE_NAME_TOKEN) > 0 Then
        Dim swTableFeat As SldWorks.Feature
        Set swTableFeat = GetFeatureFromTableAnnotation(tableAnn)
        path = Replace(path, TABLE_NAME_TOKEN, swTableFeat.Name)
    End If
    
    GetExportFilePath = GetFullPath(model, path)
    
End Function

Function GetTableData(tableAnn As SldWorks.TableAnnotation, includeHeader As Boolean) As Variant
    
    Dim tableData() As String
        
    Dim i As Integer
    Dim j As Integer
    
    Dim offset As Integer
    offset = IIf(includeHeader, 0, 1)
    
    For i = 0 + offset To tableAnn.RowCount - 1
        
        ReDim Preserve tableData(tableAnn.RowCount - 1 - offset, tableAnn.ColumnCount - 1)
        
        For j = 0 To tableAnn.ColumnCount - 1
            tableData(i - offset, j) = tableAnn.Text(i, j)
        Next
            
    Next
        
    GetTableData = tableData
    
End Function

Function FindTables(draw As SldWorks.DrawingDoc, filter As swTableAnnotationType_e, sheetName As String) As Variant
    
    Dim swTables() As SldWorks.TableAnnotation
    Dim isInit As Boolean
    isInit = False
    
    Dim vSheets As Variant
    
    vSheets = draw.GetViews()
    
    Dim i As Integer
    
    For i = 0 To UBound(vSheets)
        
        Dim vViews As Variant
        vViews = vSheets(i)
        
        Dim swSheetView As SldWorks.View
        Set swSheetView = vViews(0)
        
        If sheetName = "" Or LCase(sheetName) = LCase(swSheetView.Name) Then
        
            Dim vTableAnns As Variant
            vTableAnns = swSheetView.GetTableAnnotations
            
            If Not IsEmpty(vTableAnns) Then
                
                Dim j As Integer
                
                For j = 0 To UBound(vTableAnns)
                    
                    Dim swTableAnn As SldWorks.TableAnnotation
                    Set swTableAnn = vTableAnns(j)
                    
                    If swTableAnn.Type = filter Then
                        
                        If isInit Then
                            ReDim Preserve swTables(UBound(swTables) + 1)
                        Else
                            ReDim swTables(0)
                            isInit = True
                        End If
                        
                        Set swTables(UBound(swTables)) = swTableAnn
                        
                    End If
                    
                Next
                
            End If
        
        End If
        
    Next
    
    If isInit Then
        FindTables = swTables
    Else
        FindTables = Empty
    End If
    
End Function

Function GetSelectedTables(model As SldWorks.ModelDoc2) As Variant

    Dim swTables() As SldWorks.TableAnnotation
    Dim isInit As Boolean
    isInit = False

    Dim i As Integer
    
    Dim swSelMgr As SldWorks.SelectionMgr
    Set swSelMgr = model.SelectionManager
    
    For i = 1 To swSelMgr.GetSelectedObjectCount2(-1)
        
        Dim swSelType As Long
        swSelType = swSelMgr.GetSelectedObjectType3(i, -1)
        
        If swSelType = swSelectType_e.swSelANNOTATIONTABLES Or swSelType = swSelectType_e.swSelREVISIONTABLE Then
        
            If isInit Then
                ReDim Preserve swTables(UBound(swTables) + 1)
            Else
                ReDim swTables(0)
                isInit = True
            End If
                    
            Set swTables(UBound(swTables)) = swSelMgr.GetSelectedObject6(i, -1)
        End If
    
    Next
    
    If isInit Then
        GetSelectedTables = swTables
    Else
        GetSelectedTables = Empty
    End If
    
End Function

Sub WriteCsvFile(filePath As String, table As Variant, append As Boolean)
    
    Dim fileNmb As Integer
    fileNmb = FreeFile
    
    If append Then
        Open filePath For Append As #fileNmb
    Else
        Open filePath For Output As #fileNmb
    End If
    
    Dim i As Integer
    Dim j As Integer
    
    For i = 0 To UBound(table, 1)
        
        Dim rowContent As String
        rowContent = ""
        
        For j = 0 To UBound(table, 2)
            Dim cell As String
            cell = table(i, j)
            If HasSpecialSymbols(cell) Then
                cell = """" & ReplaceSpecialSymbols(cell) & """"
            End If
            rowContent = rowContent & IIf(j = 0, "", ",") & cell
        Next
        
        Print #fileNmb, rowContent
        
    Next
    
    Close #fileNmb
    
End Sub

Function GetFullPath(model As SldWorks.ModelDoc2, path As String)
    
    GetFullPath = path
        
    If IsPathRelative(path) Then
        
        If Left(path, 1) <> "\" Then
            path = "\" & path
        End If
        
        Dim modelPath As String
        Dim modelDir As String
        
        modelPath = model.GetPathName
        
        modelDir = Left(modelPath, InStrRev(modelPath, "\") - 1)
                
        GetFullPath = modelDir & path
    Else
        GetFullPath = path
    End If
    
End Function

Function GetFileNameWithoutExtension(path As String) As String
    GetFileNameWithoutExtension = Mid(path, InStrRev(path, "\") + 1, InStrRev(path, ".") - InStrRev(path, "\") - 1)
End Function

Function IsPathRelative(path As String)
    IsPathRelative = Mid(path, 2, 1) <> ":" And Not IsPathUnc(path)
End Function

Function IsPathUnc(path As String)
    IsPathUnc = Left(path, 2) = "\\"
End Function

Function GetFeatureFromTableAnnotation(tableAnn As SldWorks.TableAnnotation) As SldWorks.Feature
    
    Dim swTableFeat As SldWorks.Feature
    
    Select Case tableAnn.Type
                
        Case swTableAnnotationType_e.swTableAnnotation_BillOfMaterials
            
            Dim swBomTableAnn As SldWorks.BomTableAnnotation
            Set swBomTableAnn = tableAnn
            Set swTableFeat = swBomTableAnn.BomFeature.GetFeature()
            
        Case swTableAnnotationType_e.swTableAnnotation_General
            
            Dim swGenTableAnn As SldWorks.GeneralTableAnnotation
            Set swGenTableAnn = tableAnn
            Set swTableFeat = swGenTableAnn.GeneralTable.GetFeature()
        
        Case swTableAnnotationType_e.swTableAnnotation_WeldmentCutList
            
            Dim swWeldCutListTableAnn As SldWorks.WeldmentCutListAnnotation
            Set swWeldCutListTableAnn = tableAnn
            Set swTableFeat = swWeldCutListTableAnn.WeldmentCutListFeature.GetFeature()
            
        Case swTableAnnotationType_e.swTableAnnotation_BendTable
            
            Dim swBendTableAnn As SldWorks.BendTableAnnotation
            Set swBendTableAnn = tableAnn
            Set swTableFeat = swBendTableAnn.BendTable.GetFeature()
            
        Case swTableAnnotationType_e.swTableAnnotation_GeneralTolerance
            
            Dim swGeneralToleranceTableAnn As SldWorks.GeneralToleranceTableAnnotation
            Set swGeneralToleranceTableAnn = tableAnn
            Set swTableFeat = swGeneralToleranceTableAnn.GeneralToleranceTableFeature.GetFeature()
            
        Case swTableAnnotationType_e.swTableAnnotation_HoleChart
            
            Dim swHoleTableAnn As SldWorks.HoleTableAnnotation
            Set swHoleTableAnn = tableAnn
            Set swTableFeat = swHoleTableAnn.HoleTable.GetFeature()
            
        Case swTableAnnotationType_e.swTableAnnotation_PunchTable
        
            Dim swPunchTableAnn As SldWorks.PunchTableAnnotation
            Set swPunchTableAnn = tableAnn
            Set swTableFeat = swPunchTableAnn.PunchTable.GetFeature()
            
        Case swTableAnnotationType_e.swTableAnnotation_RevisionBlock
            
            Dim swRevisionTableAnn As SldWorks.RevisionTableAnnotation
            Set swRevisionTableAnn = tableAnn
            Set swTableFeat = swRevisionTableAnn.RevisionTableFeature.GetFeature()
            
        Case swTableAnnotationType_e.swTableAnnotation_TitleBlock
        
            Dim swTitleBlockTableAnn As SldWorks.TitleBlockTableAnnotation
            Set swTitleBlockTableAnn = tableAnn
            Set swTableFeat = swTitleBlockTableAnn.TitleBlockTableFeature.GetFeature()
            
        Case swTableAnnotationType_e.swTableAnnotation_WeldTable
        
            Dim swWeldTableAnn As SldWorks.WeldmentCutListAnnotation
            Set swWeldTableAnn = tableAnn
            Set swTableFeat = swWeldTableAnn.WeldmentCutListFeature.GetFeature()
        
    End Select
    
    Set GetFeatureFromTableAnnotation = swTableFeat
    
End Function

Function GetSheetFromTableAnnotation(draw As SldWorks.DrawingDoc, tableAnn As SldWorks.TableAnnotation) As SldWorks.Sheet

    Dim vSheets As Variant
    
    vSheets = draw.GetViews()
    
    Dim i As Integer
    
    For i = 0 To UBound(vSheets)
        
        Dim vViews As Variant
        vViews = vSheets(i)
        
        Dim swSheetView As SldWorks.View
        Set swSheetView = vViews(0)
        
        Dim vTableAnns As Variant
        vTableAnns = swSheetView.GetTableAnnotations
        
        If Not IsEmpty(vTableAnns) Then
            
            Dim j As Integer
            
            For j = 0 To UBound(vTableAnns)
                
                Dim swTableAnn As SldWorks.TableAnnotation
                Set swTableAnn = vTableAnns(j)
                
                If swTableAnn Is tableAnn Then
                    
                    Dim swSheet As SldWorks.Sheet
                    Set swSheet = draw.Sheet(swSheetView.GetName2())
                    Set GetSheetFromTableAnnotation = swSheet
                    Exit Function
                    
                End If
                
            Next
            
        End If
        
    Next
    
    Err.Raise vbError, "", "Table does not belong to sheet"

End Function

Function HasSpecialSymbols(cell As String) As Boolean
    HasSpecialSymbols = InStr(cell, ",") > 0 Or InStr(cell, vbLf) > 0 Or InStr(cell, vbNewLine) > 0 Or InStr(cell, """") > 0
End Function

Function ReplaceSpecialSymbols(cell As String) As String
    cell = Replace(cell, """", """""")
    ReplaceSpecialSymbols = cell
End Function