Enum MyEnum_e
    Val1 'automatically assigned value 0
    Val2 = 5 'explicitly assigned value 5
    Val3 'next automatically assigned number 6
End Enum

Enum MyIncrementEnum_e
    Val1 '0
    Val2 = Val1 + 3 '3
    Val3 = Val2 + 4 '7
End Enum

Sub main()
    
    '0 5 6
    Debug.Print MyEnum_e.Val1 & " " & MyEnum_e.Val2 & " " & MyEnum_e.Val3
    
    '0 3 7
    Debug.Print MyIncrementEnum_e.Val1 & " " & MyIncrementEnum_e.Val2 & " " & MyIncrementEnum_e.Val3
    
    'assigning the value to the variable
    Dim val As MyEnum_e
    val = MyEnum_e.Val2
    
End Sub