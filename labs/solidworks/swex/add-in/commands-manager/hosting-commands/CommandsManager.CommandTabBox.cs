using CodeStack.SwEx.AddIn.Attributes;
using CodeStack.SwEx.AddIn.Enums;
using SolidWorks.Interop.swconst;

public enum CommandsC_e
{
    [CommandItemInfo(true, true, swWorkspaceTypes_e.Assembly,
        true, swCommandTabButtonTextDisplay_e.swCommandTabButton_NoText)]
    CommandC1,

    [CommandItemInfo(true, true, swWorkspaceTypes_e.AllDocuments,
        true, swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow)]
    CommandC2,

    [CommandItemInfo(true, true, swWorkspaceTypes_e.AllDocuments,
        true, swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal)]
    CommandC3,
}