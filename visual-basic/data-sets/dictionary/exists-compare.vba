Sub ExistsCompareMode()
    
    Dim dict As Dictionary
    
    Set dict = New Dictionary
    
    dict.Add "A", 1
    dict.Add "B", 2
    dict.Add "C", 3
    dict.Add "D", 4
    
    'False
    Debug.Print dict.Exists("a")
    
    dict.Add "d", 5 'allows to add the element as the default comparison is binary
    
    'dict.CompareMode = TextCompare 'Run-time error 5: Invalid procedure call or argument
    
    Dim dict1 As New Dictionary
    dict1.CompareMode = TextCompare 'case-insensitive comparison
    
    dict1.Add "A", 1
    dict1.Add "B", 2
    dict1.Add "a", 3 'Run-time error 457: This key is already associated with an element of this collection
    
    'True
    Debug.Print dict1.Exists("a")
    
End Sub