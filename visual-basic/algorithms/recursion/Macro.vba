Sub main()
    
    Dim bom As New BomItem
    bom.Name = "A"
    bom.Qty = 1
    
    Dim bomChildren(1) As BomItem
        
    Set bomChildren(0) = New BomItem
    bomChildren(0).Name = "B"
    bomChildren(0).Qty = 2
    
    Set bomChildren(1) = New BomItem
    bomChildren(1).Name = "C"
    bomChildren(1).Qty = 3
    
    bom.Children = bomChildren
    
    Dim bomSubChildren(2) As BomItem
        
    Set bomSubChildren(0) = New BomItem
    bomSubChildren(0).Name = "D"
    bomSubChildren(0).Qty = 1
    
    Set bomSubChildren(1) = New BomItem
    bomSubChildren(1).Name = "E"
    bomSubChildren(1).Qty = 5
    
    Set bomSubChildren(2) = New BomItem
    bomSubChildren(2).Name = "F"
    bomSubChildren(2).Qty = 1
    
    bomChildren(0).Children = bomSubChildren
    
    PrintBom bom
    
End Sub

'--- recursion
Sub PrintBom(bom As BomItem, Optional level As Integer = 0)
    
    Dim offset As String
    offset = String(level, "-")
    
    Debug.Print offset & bom.Name & " (" & bom.Qty & ")"
    
    If Not IsEmpty(bom.Children) Then
        Dim i As Integer
        For i = 0 To UBound(bom.Children)
            Dim child As BomItem
            Set child = bom.Children(i)
            PrintBom child, level + 1
        Next
    End If
    
End Sub
'---