Sub AddItemsWithKeys()

    Dim mathConstColl As Collection
    Set mathConstColl = New Collection
    
    mathConstColl.Add 3.14, "PI" 'number PI
    mathConstColl.Add 9.8, "G" 'gravitational constant
    mathConstColl.Add 2.71, "e" 'Euler's number
    
    Dim i As Integer
    
    'traverse all
    For i = 1 To mathConstColl.Count()
        Debug.Print mathConstColl(i) 'item can be accessed directly without the Item property
    Next
    
    'Access values by key
    Debug.Print mathConstColl("PI")
    Debug.Print mathConstColl("G")
    Debug.Print mathConstColl("e")
    
    mathConstColl.Remove 1
    mathConstColl.Remove "e"

End Sub