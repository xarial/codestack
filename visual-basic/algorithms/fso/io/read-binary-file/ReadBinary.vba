Function ReadByteArrFromFile(filePath) As Byte()

    Dim buff() As Byte
    
    Dim fileNumb As Integer
    fileNumb = FreeFile
    
    Open filePath For Binary Access Read As fileNumb
    
    ReDim buff(0 To LOF(fileNumb) - 1)
    
    Get fileNumb, , buff
    
    Close fileNumb
    
    ReadByteArrFromFile = buff
    
End Function