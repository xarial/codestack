[Title(typeof(Resources), nameof(Resources.LocalizedPmPageTitle))]
public class LocalizedPmPage
{
    [Title(typeof(Resources), nameof(Resources.TextFieldTitle))]
    [Summary(typeof(Resources), nameof(Resources.TextFieldDescription))]
    public string TextField { get; set; }

    [Title(typeof(Resources), nameof(Resources.NumericFieldTitle))]
    [Summary(typeof(Resources), nameof(Resources.NumericFieldDescription))]
    public double NumericField { get; set; }
}