public class SelectionBoxListDataModel
{
    [SelectionBox(swSelectType_e.swSelSOLIDBODIES)]
    public List<IBody2> Bodies { get; set; } = new List<IBody2>();

    [SelectionBox(swSelectType_e.swSelEDGES, swSelectType_e.swSelNOTES, swSelectType_e.swSelCOORDSYS)]
    public List<object> Dispatches { get; set; } = new List<object>();
}