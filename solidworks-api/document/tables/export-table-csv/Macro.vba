Const OUT_FILE_PATH As String = "D:\bom.csv" 'empty string to save in the model's folder
Const INCLUDE_HEADER As Boolean = True

Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swModel As SldWorks.ModelDoc2
    
    Set swModel = swApp.ActiveDoc
    
    If Not swModel Is Nothing Then
        
        Dim swTableAnn As SldWorks.TableAnnotation
        Set swTableAnn = swModel.SelectionManager.GetSelectedObject6(1, -1)
        
        Dim vTable As Variant
        vTable = GetTableData(swTableAnn, INCLUDE_HEADER)
            
        WriteCsvFile GetExportFilePath(swModel), vTable
        
    Else
        MsgBox "Please open document"
    End If
    
End Sub

Function GetExportFilePath(model As SldWorks.ModelDoc2) As String
    
    If OUT_FILE_PATH = "" Then
        
        Dim modelPath As String
        modelPath = model.GetPathName
        
        If modelPath = "" Then
            Err.Raise vbError, "", "Model is not saved. Specify the full path to save a table or save the model"
        End If
        
        GetExportFilePath = Left(modelPath, InStrRev(modelPath, ".")) + "csv"
    Else
        GetExportFilePath = OUT_FILE_PATH
    End If
    
End Function

Function GetTableData(tableAnn As SldWorks.TableAnnotation, includeHeader As Boolean) As Variant
    
    Dim tableData() As String
        
    Dim i As Integer
    Dim j As Integer
    
    Dim offset As Integer
    offset = IIf(INCLUDE_HEADER, 0, 1)
    
    For i = 0 + offset To tableAnn.RowCount - 1
        
        ReDim Preserve tableData(tableAnn.RowCount - 1 - offset, tableAnn.ColumnCount - 1)
        
        For j = 0 To tableAnn.ColumnCount - 1
            tableData(i - offset, j) = tableAnn.Text(i, j)
        Next
            
    Next
        
    GetTableData = tableData
    
End Function

Sub WriteCsvFile(filePath As String, table As Variant)
    
    Dim fileNmb As Integer
    fileNmb = FreeFile
    
    Open filePath For Output As #fileNmb
    
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

Function HasSpecialSymbols(cell As String) As Boolean
    
    HasSpecialSymbols = InStr(cell, ",") > 0 Or InStr(cell, vbLf) > 0 Or InStr(cell, vbNewLine) > 0 Or InStr(cell, """") > 0
    
End Function

Function ReplaceSpecialSymbols(cell As String) As String
    cell = Replace(cell, """", """""")
    ReplaceSpecialSymbols = cell
End Function