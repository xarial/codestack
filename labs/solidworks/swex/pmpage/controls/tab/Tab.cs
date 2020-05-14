using CodeStack.SwEx.Common.Attributes;
using CodeStack.SwEx.PMPage.Attributes;
using CodeStack.SwEx.Properties;

public class TabDataModel
{
    [Tab]
    [Icon(typeof(Resources), nameof(Resources.OffsetImage))]
    public class TabControl1
    {
        public string Field1 { get; set; }
    }

    public TabControl1 Tab1 { get; set; }

}
