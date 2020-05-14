//...
private IDocumentsHandler<DocumentHandler> m_DocHandler;
//...
    //...
    m_DocHandler = CreateDocumentsHandler();
    m_DocHandler.HandlerCreated += OnDocHandlerCreated;
    //...
private void OnDocHandlerCreated(DocumentHandler doc)
{
    doc.Rebuild += OnDocRebuild;
    doc.Save += OnDocSave;

}

private bool OnDocRebuild(DocumentHandler docHandler, RebuildState_e state)
{
    //TODO: handle rebuild
    return true;
}

private bool OnDocSave(DocumentHandler docHandler, string fileName, SaveState_e state)
{
    //TODO: handle saving
    return true;
}