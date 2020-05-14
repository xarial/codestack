<Title(GetType(Resources), NameOf(Resources.ToolbarTitle)), Description("Toolbar with commands")>
<Icon(GetType(Resources), NameOf(Resources.commands))>
Public Enum Commands_e
    <Title("Command 1"), Description("Sample command 1")>
    <Icon(GetType(Resources), NameOf(Resources.command1))>
    <CommandItemInfo(True, True, swWorkspaceTypes_e.Assembly, True, swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow)>
    Command1
    Command2
End Enum
    '...
    AddCommandGroup(New Action(Of Commands_e)(AddressOf OnButtonClick))
    '...
Private Sub OnButtonClick(ByVal cmd As Commands_e)
End Sub