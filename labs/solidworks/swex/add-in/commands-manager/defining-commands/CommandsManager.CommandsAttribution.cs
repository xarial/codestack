using CodeStack.SwEx.Common.Attributes;
using CodeStack.SwEx.Properties;
using System.ComponentModel;

[Title(typeof(Resources), nameof(Resources.ToolbarTitle))]
[Description("Command Group Title")]
[Icon(typeof(Resources), nameof(Resources.commands))]
public enum CommandsB_e
{
    [Title("First Command")]
    [Description("Hint text for first command")]
    [Icon(typeof(Resources), nameof(Resources.command1))]
    CommandB1,

    [Title("Second Command")]
    [Description("Hint text for second command")]
    [Icon(typeof(Resources), nameof(Resources.command2))]
    CommandB2,

    [Title("Third Command")]
    [Description("Hint text for third command")]
    [Icon(typeof(Resources), nameof(Resources.command3))]
    CommandB3
}