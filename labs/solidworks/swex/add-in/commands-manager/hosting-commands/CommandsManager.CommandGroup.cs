public override bool OnConnect()
{
    AddCommandGroup<CommandsA_e>(OnCommandsAButtonClick);
    AddCommandGroup<CommandsB_e>(OnCommandsBButtonClick);
    AddCommandGroup<CommandsC_e>(OnCommandsCButtonClick);

    return true;
}

private void OnCommandsAButtonClick(CommandsA_e cmd)
{
    //TODO: handle the button click
}

private void OnCommandsBButtonClick(CommandsB_e cmd)
{
    //TODO: handle the button click
}

private void OnCommandsCButtonClick(CommandsC_e cmd)
{
    //TODO: handle the button click
}