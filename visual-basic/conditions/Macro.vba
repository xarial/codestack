Sub main()

    Dim myVar As Integer
    myVar = 25
    
    '--- fallback
    If myVar > 10 Then
        Debug.Print "Value of myVar variable is greater than 10"
    Else
        Debug.Print "Value of myVar variable is lower than 10"
    End If
    '---
    
    '--- multiple
    If myVar < 0 Then
        Debug.Print "myVar has a negative value"
    ElseIf myVar = 0 Then
        Debug.Print "myVar equals to 0"
    ElseIf myVar > 0 And myVar < 10 Then
        Debug.Print "myVar value in a range of 0...10 (exclusive)"
    Else
        Debug.Print "Value of myVar is 10 or more"
    End If
    '---
    
End Sub

Sub main2()

    Dim dayOfTheWeek As Integer
    dayOfTheWeek = 3
    
    '--- select-case
    Select Case dayOfTheWeek
        Case 1
            Debug.Print "Monday"
        Case 2
            Debug.Print "Tuesday"
        Case 3
            Debug.Print "Wednesday"
        Case 4
            Debug.Print "Thursday"
        Case 5
            Debug.Print "Friday"
        Case 6
            Debug.Print "Saturday"
        Case 7
            Debug.Print "Sunday"
        Case Else
            Err.Raise vbError, "", "Value outside of the 1...7 range"
    End Select
    '---

End Sub

Sub main3()
    
    '--- logical
    Dim varA, varB, varC, varD As Boolean
        
    varA = True
    varB = False
    varC = True
    varD = False
    
    Debug.Print varA And varB 'False
    Debug.Print Not (varA And varB) 'True
    Debug.Print varA And varC 'True
    Debug.Print varA Or varC 'True
    Debug.Print varA Or varB 'True
    Debug.Print varB Or varD 'False
    Debug.Print (varA Or varB) And varD 'False
    Debug.Print varA Or (varB And varD) 'True
    '---
    
End Sub