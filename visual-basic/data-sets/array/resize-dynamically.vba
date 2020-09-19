Dim evenNumbersArr() As Integer

Dim i As Integer

For i = 0 To 100
    If i Mod 2 = 0 Then
                
        If (Not evenNumbersArr) = -1 Then
            ReDim evenNumbersArr(0)
        Else
            ReDim Preserve evenNumbersArr(UBound(evenNumbersArr) + 1)
        End If
        
        evenNumbersArr(UBound(evenNumbersArr)) = i
    End If
Next