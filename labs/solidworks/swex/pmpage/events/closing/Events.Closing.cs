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

    m_Page.Handler.Closing += OnClosing;

    return true;
}

private void OnClosing(swPropertyManagerPageCloseReasons_e reason, ClosingArg arg)
{
    if (reason == swPropertyManagerPageCloseReasons_e.swPropertyManagerPageClose_Okay)
    {
        if (string.IsNullOrEmpty(m_Data.Text))
        {
            arg.Cancel = true;
            arg.ErrorTitle = "Insert Note Error";
            arg.ErrorMessage = "Please specify the note text";
        }
    }
}