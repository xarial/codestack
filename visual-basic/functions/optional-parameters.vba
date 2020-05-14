Sub main()
    
    Debug.Print Pow(2) '4
    Debug.Print Pow(2, 3) '8

    PrintAddress state:="NSW", postcode:=2000 'Australia NSW 2000
    
End Sub

Function Pow(number As Double, Optional power As Double = 2) As Double
    
    Pow = number ^ power
    
End Function

Sub PrintAddress(Optional country As String = "Australia", Optional state As String = "", Optional suburb As String = "", Optional postcode As Integer = 0, Optional streetName As String = "", Optional buildingNumber As Integer = 0, Optional unitNumber As Integer = 0)

    If country <> "" Then
        Debug.Print country
    End If
    
    If state <> "" Then
        Debug.Print state
    End If
    
    If suburb <> "" Then
        Debug.Print suburb
    End If
    
    If postcode > 0 Then
        Debug.Print postcode
    End If
    
    If streetName <> "" Then
        Debug.Print streetName
    End If
    
    If buildingNumber > 0 Then
        Debug.Print buildingNumber
    End If
    
    If unitNumber > 0 Then
        Debug.Print "Unit: " & unitNumber
    End If
    
End Sub