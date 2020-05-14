public class SelectionBoxDataModel
{
    [SelectionBox(swSelectType_e.swSelSOLIDBODIES)]
    public IBody2 Body { get; set; }

    [SelectionBox(swSelectType_e.swSelEDGES, swSelectType_e.swSelNOTES, swSelectType_e.swSelCOORDSYS)]
    public object Dispatch { get; set; }
}