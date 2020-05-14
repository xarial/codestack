Sub main()
    
    'prints ProcedureWithoutParameters twice
    ProcedureWithoutParameters
    ProcedureWithoutParameters
    
    'Compile error: Argument not optional
    'SayHello
    
    'Hello, Test
    SayHello "Test"
    
    Dim formDate As String
    FormatDate "dd-MM-yyyy", formDate
    
    '20-06-2018
    Debug.Print formDate
    
    '20-06-2018
    Debug.Print GetFormattedDate("dd-MM-yyyy")
    
End Sub

Sub ProcedureWithoutParameters()
    
    Debug.Print "ProcedureWithoutParameters"

End Sub

Sub SayHello(name As String)
    
    Debug.Print "Hello, " & name

End Sub

Sub FormatDate(dateFormat As String, ByRef formattedDate As String)
    
    Dim curDate As Date
    curDate = Now
    
    formattedDate = format(curDate, dateFormat)
    
End Sub

Function GetFormattedDate(dateFormat As String) As String
    
    Dim curDate As Date
    curDate = Now
    
    GetFormattedDate = format(curDate, dateFormat)
    
End Function