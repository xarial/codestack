Property Get Prop1() As Integer
    Prop1 = Prop1
End Property

Property Let Prop1(val As String)
    Debug.Print val
End Property

Property Get Prop2()
    Prop2 = "Prop2 Value"
End Property

Property Let Prop2(val As String) 'invalid as implicitly assigned type is Variant
    Debug.Print val
End Property