[Title("AddInEx Commands")]
public enum Commands_e
{
    Command1,

    [CommandSpacer]
    Command2
}

[Title("Sub Menu Commands")]
[CommandGroupInfo(typeof(Commands_e))]
public enum SubCommands_e
{
    SubCommand1,
    SubCommand2
}

public override bool OnConnect()
{
    AddCommandGroup<Commands_e>(OnButtonClick);
    AddCommandGroup<SubCommands_e>(OnButtonClick);
    return true;
}

private void OnButtonClick(Commands_e cmd)
{
}

private void OnButtonClick(SubCommands_e cmd)
{
}