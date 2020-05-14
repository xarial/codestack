public class DataModel
{
    public string Text { get; set; }
}

private DataModel m_Data;

private PropertyManagerPageEx<MyPMPageHandler, DataModel> m_Page;

public override bool OnConnect()
{
    m_Data = new DataModel();
    m_Page = new PropertyManagerPageEx<MyPMPageHandler, DataModel>(App);

    m_Page.Handler.DataChanged += OnDataChanged;

    return true;
}

private void OnDataChanged()
{
    var text = m_Data.Text;
    //TODO: handle the data changing, e.g. update preview
}