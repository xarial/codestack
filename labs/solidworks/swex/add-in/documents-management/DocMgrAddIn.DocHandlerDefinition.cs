public class MyDocHandler : IDocumentHandler
{
    private IModelDoc2 m_Model;

    public void Init(ISldWorks app, IModelDoc2 model)
    {
        if (model is PartDoc)
        {
            m_Model = model;
            (m_Model as PartDoc).AddItemNotify += OnAddItemNotify;
        }
        //TODO: handle other doc types
    }

    private int OnAddItemNotify(int EntityType, string itemName)
    {
        //Implement
        return 0;
    }

    public void Dispose()
    {
        if (m_Model is PartDoc)
        {
            (m_Model as PartDoc).AddItemNotify -= OnAddItemNotify;
        }
    }
}