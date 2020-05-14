public class DataModelEnable
{
    [ControlTag(nameof(Enable))]
    public bool Enable { get; set; }

    [SelectionBox(swSelectType_e.swSelFACES)]
    [DependentOn(typeof(EnableDepHandler), nameof(Enable))]
    public IEntity Selection { get; set; }
}

public class EnableDepHandler : DependencyHandler
{
    protected override void UpdateControlState(IPropertyManagerPageControlEx control, IPropertyManagerPageControlEx[] parents)
    {
        control.Enabled = (bool)parents.First().GetValue();
    }
}