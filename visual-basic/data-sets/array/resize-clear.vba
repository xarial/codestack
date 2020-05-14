Sub ResizeAndClearArray()
    
    Dim doubleArr() As Double
    Dim i As Integer
    
    ReDim doubleArr(2)
    
    For i = 0 To UBound(doubleArr)
        doubleArr(i) = i + 1
    Next
    
    'resizing and clearing the array
    ReDim doubleArr(3)
    doubleArr(3) = 4
    
    '0 0 0 4
    For i = 0 To UBound(doubleArr)
        Debug.Print doubleArr(i)
    Next

End Sub