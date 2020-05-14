public override bool OnConnect()
{
    AddContextMenu<CommandsD_e>(OnCommandsDContextMenuClick);
    AddContextMenu<CommandsE_e>(OnCommandsEContextMenuClick, swSelectType_e.swSelFACES);

    return true;
}

private void OnCommandsDContextMenuClick(CommandsD_e cmd)
{
    //TODO: handle the context menu click
}

private void OnCommandsEContextMenuClick(CommandsE_e cmd)
{
    //TODO: handle the context menu click
}