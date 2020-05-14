private IDocumentsHandler<MyDocHandler> m_DocHandler;
private IDocumentsHandler<DocumentHandler> m_DocHandlerGeneric;

public override bool OnConnect()
{
    m_DocHandler = CreateDocumentsHandler<MyDocHandler>();
    m_DocHandlerGeneric = CreateDocumentsHandler();
    m_DocHandlerGeneric.HandlerCreated += OnHandlerCreated;
    return true;
}

private void OnHandlerCreated(DocumentHandler doc)
{
    //TODO: implement
}