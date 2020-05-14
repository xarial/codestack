Dim WithEvents swPart As SldWorks.PartDoc

Private Function swPart_AddItemNotify(ByVal EntityType As Long, ByVal itemName As String) As Long

    If EntityType = swNotifyEntityType_e.swNotifyFeature Then
        MsgBox itemName & " feature is added"
    End If
    
End Function

Sub SetPart(part As SldWorks.PartDoc)
    
    Set swPart = part
    
End Sub