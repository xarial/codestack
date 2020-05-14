using CodeStack.SwEx.AddIn.Attributes;
using CodeStack.SwEx.AddIn.Enums;

public enum CommandsD_e
{
    [CommandItemInfo(swWorkspaceTypes_e.Part)]
    CommandD1,

    [CommandItemInfo(swWorkspaceTypes_e.Part | swWorkspaceTypes_e.Assembly)]
    CommandD2
}