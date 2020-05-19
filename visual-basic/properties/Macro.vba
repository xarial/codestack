'--- read-only
Sub Macro1()
    Debug.Print Prop1
    Debug.Print Prop2 Is Nothing
    'Prop1.Value = "New Value" - not possible
End Sub

Property Get Prop1() As String
    Prop1 = "Prop1 Value"
End Property

Property Get Prop2() As Object
    Set Prop2 = CreateObject("Scripting.Dictionary")
End Property
'---

'--- write-only
Sub Macro2()
    Prop3 = "Val1"
    Set Prop4 = CreateObject("Scripting.Dictionary")
End Sub

Property Let Prop3(val As String)
    Debug.Print val
End Property

Property Set Prop4(val As Object)
    Debug.Print val Is Nothing
End Property
'---

'--- read-write
Sub Macro3()
    Prop5 = "Val1"
    Debug.Print Prop5
End Sub

Property Get Prop5() As String
    Prop5 = "Prop5 Value"
End Property

Property Let Prop5(val As String)
    Debug.Print val
End Property
'---

'--- parameters
Sub Macro4()
    Prop6("p1") = 20
    Debug.Print Prop6("p2")
End Sub

Property Get Prop6(param1 As String) As Double
    Prop6 = 10
End Property

Property Let Prop6(param1 As String, val As Double)
    Debug.Print param1 & " " & val
End Property
'---