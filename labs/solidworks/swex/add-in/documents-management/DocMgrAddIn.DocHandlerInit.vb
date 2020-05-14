Private m_DocHandler As IDocumentsHandler(Of MyDocHandler)
Private m_DocHandlerGeneric As IDocumentsHandler(Of DocumentHandler)

Public Overrides Function OnConnect() As Boolean
    m_DocHandler = CreateDocumentsHandler(Of MyDocHandler)()
    m_DocHandlerGeneric = CreateDocumentsHandler()
    AddHandler m_DocHandlerGeneric.HandlerCreated, AddressOf OnHandlerCreated
    Return True
End Function

Private Sub OnHandlerCreated(ByVal doc As DocumentHandler)
    'TODO: implement
End Sub