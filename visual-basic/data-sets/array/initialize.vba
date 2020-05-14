Sub InitializeArray()
    
    Dim doubleArr() As Double 'not initialized array
    
    'Array is initialized = False
    Debug.Print "Array is initialized = " & IsArrayInitialized(doubleArr)
    
    ReDim doubleArr(2) 'resizing array to hold 3 doubles
    
    'Array is initialized = True of size 3
    Debug.Print "Array is initialized = " & IsArrayInitialized(doubleArr) & " of size " & GetArraySize(doubleArr)
    
    Dim textArr(4) As String 'initialized at declaration
    'Array is initialized = True of size 5
    Debug.Print "Array is initialized = " & IsArrayInitialized(textArr) & " of size " & GetArraySize(textArr)
    
    'initializing with custom boundaries
    Dim intArr(1 To 5) As Integer
    'Array is initialized = True of size 5 (1 to 5)
    Debug.Print "Array is initialized = " & IsArrayInitialized(intArr) & " of size " & GetArraySize(intArr) & " (" & LBound(intArr) & " to " & UBound(intArr) & ")"
    
    'Debug.Print intArr(0) 'Run-time error 9: subscript out of range
    
End Sub

Function IsArrayInitialized(vArr As Variant) As Boolean

    If IsArray(vArr) Then
        
        On Error GoTo End_
        
        If UBound(vArr) >= 0 Then
            IsArrayInitialized = True
            Exit Function
        End If
        
    End If

End_:

    IsArrayInitialized = False
    
End Function

Function GetArraySize(vArr As Variant) As Integer
    
    If IsArrayInitialized(vArr) Then
        GetArraySize = UBound(vArr) - LBound(vArr) + 1
    Else
        GetArraySize = 0
    End If
    
End Function