'...
Private m_DocHandler As IDocumentsHandler(Of DocumentHandler)
'...
    '...
    m_DocHandler = CreateDocumentsHandler()
    AddHandler m_DocHandler.HandlerCreated, AddressOf OnDocHandlerCreated
    '...
Private Sub OnDocHandlerCreated(ByVal doc As DocumentHandler)
    AddHandler doc.Rebuild, AddressOf OnDocRebuild
    AddHandler doc.Save, AddressOf OnDocSave
End Sub

Private Function OnDocRebuild(ByVal docHandler As DocumentHandler, ByVal state As RebuildState_e) As Boolean
    'TODO: handle rebuild
    Return True
End Function

Private Function OnDocSave(ByVal docHandler As DocumentHandler, ByVal fileName As String, ByVal state As SaveState_e) As Boolean
    'TODO: handle saving
    Return True
End Function