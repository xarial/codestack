Type MyType
    IntValue As Integer
    DoubleValue As Double
    StringValue As String
End Type

Sub main()

    Dim val1 As MyType
    val1.DoubleValue = 10.5
    val1.IntValue = 5
    val1.StringValue = "Hello World"
    
    Dim val2 As MyType
    val2 = val1 'all values are copied
    
    val2.DoubleValue = 2.5
    val2.StringValue = "Modified Hello World"
    val2.IntValue = 1
    
    '10.5, 5, Hello World
    Debug.Print val1.DoubleValue & ", " & val1.IntValue & ", " & val1.StringValue
    
    '2.5, 1, Modified Hello World
    Debug.Print val2.DoubleValue & ", " & val2.IntValue & ", " & val2.StringValue
    
End Sub