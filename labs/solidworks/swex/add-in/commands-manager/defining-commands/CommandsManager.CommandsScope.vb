Imports CodeStack.SwEx.AddIn.Attributes
Imports CodeStack.SwEx.AddIn.Enums

Public Enum CommandsD_e

    <CommandItemInfo(swWorkspaceTypes_e.Part)>
    CommandD1

    <CommandItemInfo(swWorkspaceTypes_e.Part Or swWorkspaceTypes_e.Assembly)>
    CommandD2

End Enum
