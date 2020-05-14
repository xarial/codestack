Function ReadText(filePath As String) As String
    
    Dim fileNo As Integer

    fileNo = FreeFile
    
    Dim content As String
    
    Dim isFirstLine As Integer
    isFirstLine = True
    
    Open filePath For Input As #fileNo
    
    Do While Not EOF(fileNo)
        
        Dim line As String
        
        Line Input #fileNo, line
        
        content = content & IIf(Not isFirstLine, vbLf, "") & line
        isFirstLine = False
        
    Loop
    
    Close #fileNo
    
    ReadText = content
    
End Function