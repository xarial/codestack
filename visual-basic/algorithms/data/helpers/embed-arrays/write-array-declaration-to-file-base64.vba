Sub WriteByteArrayDeclarationToFileAsBase64(buffer As Variant, filePath As String, Optional lineMaxLength As Integer = -1)
    
    Const FUNC_NAME = "GetBufferPart"
    
    Dim fileNo As Integer
    fileNo = FreeFile
    
    Open filePath For Output As #fileNo
        
    Dim data As String
    data = ConvertToBase64String(buffer)
    data = Replace(data, vbLf, "")
    
    If lineMaxLength > 1 Then
            
        Const MAX_LINE_CONTINUATIONS As Integer = 24
        
        Dim curLineIndex As Integer
        Dim curCont As Integer
        curLineIndex = 0
        
        Dim i As Long
        
        Dim funcsCount As Integer
        funcsCount = Round((Len(data) - 1) / lineMaxLength / MAX_LINE_CONTINUATIONS) - 1
        
        Print #fileNo, "Function GetBase64Buffer() As String"
                
        For i = 0 To funcsCount
            Print #fileNo, "GetBase64Buffer = GetBase64Buffer & " & FUNC_NAME & i & "()"
        Next
        
        Print #fileNo, "End Function"
        
        Dim funcName As String
        
        For i = 1 To Len(data) Step lineMaxLength
            
            If curCont = MAX_LINE_CONTINUATIONS Then
                curCont = 0
                curLineIndex = curLineIndex + 1
            End If
            
            Dim length As Integer
        
            Dim isLast As Boolean
            isLast = False
            
            If i + lineMaxLength > Len(data) Then
                length = Len(data) - i + 1
                isLast = True
            Else
                length = lineMaxLength
            End If
            
            curCont = curCont + 1
            
            If curCont = 1 Then
                funcName = FUNC_NAME & curLineIndex
                Print #fileNo, "Function " & funcName & "() As String"
            End If
            
            isLast = isLast Or curCont >= MAX_LINE_CONTINUATIONS
            
            Dim lineConc As String
            lineConc = ""
            If Not isLast Then
                lineConc = " & _"
            End If
            
            Print #fileNo, IIf(curCont = 1, funcName & " = ", ""); """" & Mid(data, i, length) & """" & lineConc
            
            If isLast Then
                Print #fileNo, "End Function"
            End If
            
        Next
        
    Else
        Print #fileNo, data
    End If
    
    Close #fileNo
    
End Sub

Function ConvertToBase64String(vArr As Variant) As String
    
    Dim xmlDoc As Object
    Dim xmlNode As Object
    
    Set xmlDoc = CreateObject("MSXML2.DOMDocument")
    
    Set xmlNode = xmlDoc.createElement("b64")
    
    xmlNode.DataType = "bin.base64"
    xmlNode.nodeTypedValue = vArr
    
    ConvertToBase64String = xmlNode.Text
    
End Function