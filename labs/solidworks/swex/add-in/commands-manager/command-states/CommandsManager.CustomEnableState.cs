public enum Commands_e
{
    Command1,
    Command2
}

public override bool OnConnect()
{
    AddCommandGroup<Commands_e>(OnButtonClick, OnButtonEnable);
    return true;
}

private void OnButtonEnable(Commands_e cmd, ref CommandItemEnableState_e state)
{
    switch (cmd)
    {
        case Commands_e.Command1:
        case Commands_e.Command2:
            //TODO: implement logic to identify the state of the button
            state = CommandItemEnableState_e.DeselectDisable;
            break;
    }
}