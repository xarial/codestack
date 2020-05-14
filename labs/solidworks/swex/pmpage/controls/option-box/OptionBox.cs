using CodeStack.SwEx.Common.Attributes;
using CodeStack.SwEx.PMPage.Attributes;

public class OptionBoxDataModel
{
    public enum Options_e
    {
        Option1,
        Option2,
        [Title("Third Option")]
        Option3
    }

    [OptionBox]
    public Options_e Options { get; set; }
}