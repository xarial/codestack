'--- init
Dim Arr(9) As String

Sub InitArr()
    Arr(0) = "A": Arr(1) = "B": Arr(2) = "C": Arr(3) = "D": Arr(4) = "E"
    Arr(5) = "F": Arr(6) = "G": Arr(7) = "H": Arr(8) = "I": Arr(9) = "J"
End Sub
'---

'--- for
Sub ForLoop()
    
    InitArr
    
    Dim i As Integer
    
    For i = 0 To 9
        Dim val As String
        val = Arr(i)
        Debug.Print val
    Next
    
End Sub
'---

'--- for-reverse
Sub ForLoopStep()
    
    InitArr
    
    Dim i As Integer
    
    For i = UBound(Arr) To 0 Step -1
        Dim val As String
        val = Arr(i)
        Debug.Print val
    Next
    
End Sub
'---

'--- while
Sub WhileLoop()
    
    InitArr
    
    Dim i As Integer
    Dim val As String
    i = 0
    
    While val <> "D"
        val = Arr(i)
        i = i + 1
        Debug.Print val
    Wend
    
End Sub
'---

'--- do
Sub DoLoop()
    
    InitArr
    
    Dim i As Integer
    Dim val As String
    i = 0
    
    Do
        val = Arr(i)
        i = i + 1
        Debug.Print val
    Loop While val <> "D"
    
End Sub
'---

'--- foreach
Sub ForEachLoop()
            
    InitArr
    
    For Each x In Arr
        Debug.Print x
    Next
    
End Sub
'---

'--- infinite
Sub InifiniteLoop()
    
    InitArr
    
    Dim i As Integer
    i = 0
    
    While i <> UBound(Arr) + 1
        Debug.Print Arr(i)
    Wend
    
End Sub
'---