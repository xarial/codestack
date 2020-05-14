Sub WriteArrayDeclarationToFile(buffer As Variant, filePath As String, varName As String, typeName As String, Optional elemsPerRow As Integer = 10)
    
    Dim fileNo As Integer
    fileNo = FreeFile
    
    Open filePath For Output As #fileNo
    
    Print #fileNo, "Dim " & varName & "(" & UBound(buffer) & ") As " & typeName
    
    Dim i As Long
    
    For i = 0 To UBound(buffer) Step elemsPerRow
        
        Dim j As Long
        Dim last As Long
        
        If i + elemsPerRow > UBound(buffer) Then
            last = UBound(buffer)
        Else
            last = i + elemsPerRow - 1
        End If
        
        Dim line As String
        line = ""
        
        For j = i To last
            Dim val As String
            val = buffer(j)
            If LCase(typeName) = "string" Then
                val = """" & val & """"
            End If
            line = IIf(line <> "", line & ": ", "") + varName & "(" & j & ")=" & val
        Next
        
        Print #fileNo, line
        
    Next
    
    Close #fileNo
    
End Sub