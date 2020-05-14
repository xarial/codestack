Imports CodeStack.SwEx.AddIn.Attributes
Imports CodeStack.SwEx.AddIn.Enums
Imports SolidWorks.Interop.swconst

Public Enum CommandsC_e

    <CommandItemInfo(True, True, swWorkspaceTypes_e.Assembly, True, swCommandTabButtonTextDisplay_e.swCommandTabButton_NoText)>
    CommandC1

    <CommandItemInfo(True, True, swWorkspaceTypes_e.AllDocuments, True, swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow)>
    CommandC2

    <CommandItemInfo(True, True, swWorkspaceTypes_e.AllDocuments, True, swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal)>
    CommandC3

End Enum
