Imports CodeStack.SwEx.Common.Attributes
Imports CodeStack.SwEx.My.Resources
Imports System.ComponentModel

<Title(GetType(Resources), NameOf(Resources.ToolbarTitle))>
<Description("Command Group Title")>
<Icon(GetType(Resources), NameOf(Resources.commands))>
Public Enum CommandsB_e

    <Title("First Command")>
    <Description("Hint text for first command")>
    <Icon(GetType(Resources), NameOf(Resources.command1))>
    CommandB1

    <Title("Second Command")>
    <Description("Hint text for second command")>
    <Icon(GetType(Resources), NameOf(Resources.command2))>
    CommandB2

    <Title("Third Command")>
    <Description("Hint text for third command")>
    <Icon(GetType(Resources), NameOf(Resources.command3))>
    CommandB3

End Enum
