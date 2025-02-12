Sub WriteCsvFile(filePath As String, table As Variant, append As Boolean)
    
    Dim tableContent As String
    
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
        
        If i > 0 Then
            tableContent = tableContent & vbNewLine
        End If
        
        tableContent = tableContent & rowContent
        
    Next
    
    Const adTypeBinary As Integer = 1
    Const adTypeText As Integer = 2
    
    Const adSaveCreateOverWrite As Integer = 2
    
    Dim fileStreamUtf8 As Object
    Set fileStreamUtf8 = CreateObject("ADODB.Stream")
    fileStreamUtf8.Type = adTypeText
    fileStreamUtf8.Charset = "utf-8"
    fileStreamUtf8.Open
    fileStreamUtf8.WriteText tableContent
    fileStreamUtf8.Position = 3 'skip BOM
        
    Dim fileStreamUtf8NoBom As Object
    Set fileStreamUtf8NoBom = CreateObject("ADODB.Stream")
    fileStreamUtf8NoBom.Type = adTypeBinary
    fileStreamUtf8NoBom.Open
    fileStreamUtf8.CopyTo fileStreamUtf8NoBom
    fileStreamUtf8NoBom.SaveToFile filePath, adSaveCreateOverWrite
    
End Sub