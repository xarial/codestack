[Title(typeof(Resources), nameof(Resources.ToolbarTitle)), Description("Toolbar with commands")]
[Icon(typeof(Resources), nameof(Resources.commands))]
public enum Commands_e
{
    [Title("Command 1"), Description("Sample command 1")]
    [Icon(typeof(Resources), nameof(Resources.command1))]
    [CommandItemInfo(true, true, swWorkspaceTypes_e.Assembly, true, swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow)]
    Command1,
    Command2
}
    //...
    AddCommandGroup<Commands_e>(OnButtonClick);
    //...
private void OnButtonClick(Commands_e cmd)
{
    //TODO: handle commands
}