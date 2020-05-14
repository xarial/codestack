Sub AddInsertItems()

    Dim coll As Collection
    Set coll = New Collection
    
    coll.Add "A"
    coll.Add "D"
    coll.Add "B", , , 1 'insert after first element
    coll.Add "C", , 3 'insert before 3rd element
    
    Dim i As Integer
    
    'A B C D
    For i = 1 To coll.Count() 'collection is 1-base indexed
        Debug.Print coll.Item(i)
    Next
    
End Sub