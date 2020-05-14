public class DataModelIgnore
{
    public string Text { get; set; }

    [IgnoreBinding]
    public int CalculatedField { get; set; } //control will not be generated for this field
}