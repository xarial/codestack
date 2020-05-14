using CodeStack.SwEx.Common.Attributes;
using CodeStack.SwEx.PMPage.Attributes;
using CodeStack.SwEx.Properties;
using SolidWorks.Interop.swconst;

public class ComboBoxDataModel
{
    public enum OptionsCustomized_e
    {
        [Title("First Option")] //static title
        Option1,

        [Title(typeof(Resources), nameof(Resources.Option2Title))] //title loaded from resources
        Option2
    }

    public OptionsCustomized_e Options2 { get; set; }
}
