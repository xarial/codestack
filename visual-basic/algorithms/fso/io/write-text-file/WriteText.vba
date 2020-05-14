Sub WriteText(filePath As String, content As String, append As Boolean)
    
    Dim fileNo As Integer
    fileNo = FreeFile
    
    If append Then
        Open filePath For Append As #fileNo
    Else
        Open filePath For Output As #fileNo
    End If
    
    Print #fileNo, content
    
    Close #fileNo
    
End Sub