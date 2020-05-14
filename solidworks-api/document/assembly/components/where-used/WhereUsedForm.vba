Dim swComps As Variant

Public Assembly As SldWorks.AssemblyDoc

Property Let Components(val As Variant)
    swComps = val
    
    Dim i As Integer
    
    For i = 0 To UBound(swComps)
        
        Dim swComp As SldWorks.Component2
        Set swComp = swComps(i)
        
        Dim name As String
        
        If swComp Is Nothing Then
            name = "<root>"
        Else
            name = swComp.Name2
        End If
        
        ReferencesList.AddItem name
    Next
    
End Property

Private Sub ReferencesList_Change()

    Dim swComp As SldWorks.Component2
    Set swComp = swComps(ReferencesList.ListIndex)
        
    If Not swComp Is Nothing Then
        swComp.Select4 False, Nothing, False
    Else
        Assembly.ClearSelection2 False
    End If
        
End Sub