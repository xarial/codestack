private PropertyManagerPageEx<MyPMPageHandler, DataModel> m_Page;
private DataModel m_Data = new DataModel();

private enum Commands_e
{
    ShowPmpPage
}

public override bool OnConnect()
{
    m_Page = new PropertyManagerPageEx<MyPMPageHandler, DataModel>(App);
    AddCommandGroup<Commands_e>(ShowPmpPage);
    return true;
}

private void ShowPmpPage(Commands_e cmd)
{
    m_Page.Handler.Closed += OnPageClosed;
    m_Page.Show(m_Data);
}

private void OnPageClosed(swPropertyManagerPageCloseReasons_e reason)
{
    Debug.Print($"Text: {m_Data.Simple.Text}");
    Debug.Print($"Size: {m_Data.Simple.Size}");
    Debug.Print($"Number: {m_Data.Simple.Number}");
}