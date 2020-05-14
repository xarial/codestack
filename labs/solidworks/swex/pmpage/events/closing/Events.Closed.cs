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

    m_Page.Handler.Closed += OnClosed;

    return true;
}

private void OnClosed(swPropertyManagerPageCloseReasons_e reason)
{
    if (reason == swPropertyManagerPageCloseReasons_e.swPropertyManagerPageClose_Okay)
    {
        //TODO: do work
    }
    else
    {
        //TODO: release resources
    }
}