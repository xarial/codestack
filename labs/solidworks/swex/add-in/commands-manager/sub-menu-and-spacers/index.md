---
title: Adding sub-menus and spacers to SOLIDWORKS command manager using SwEx.AddIn
caption: Sub-Menus And Spacers
description: Adding sub-menus and spacers or command tab boxes in SOLIDWORKS command manager using SwEx.AddIn framework
image: sub-menu-and-spacer.png
toc-group-name: labs-solidworks-swex
order: 3
---
## Adding spacer

Spacer can be added between the commands by decorating the command using the [CommandSpacerAttribute](https://docs.codestack.net/swex/add-in/html/T_CodeStack_SwEx_AddIn_Attributes_CommandSpacerAttribute.htm). Spacer will be added before this command.

{% code-snippet { file-name: CommandsManager.Spacer.* } %}

If command tab tab boxes are created for this command group (i.e. *showInCmdTabBox* parameter is set to *true* in the [CommandItemInfoAttribute](https://docs.codestack.net/swex/add-in/html/M_CodeStack_SwEx_AddIn_Attributes_CommandItemInfoAttribute__ctor_2.htm)), spacer is not reflected in the corresponding command tab box.

## Adding sub-menus

Sub-menus for the command groups can be defined by calling the corresponding overload of the [CommandGroupInfoAttribute](https://docs.codestack.net/swex/add-in/html/M_CodeStack_SwEx_AddIn_Attributes_CommandGroupInfoAttribute__ctor_2.htm) attribute and specifying the type of the parent menu group

{% code-snippet { file-name: CommandsManager.SubMenu.* } %}

Sub menus are rendered in separate tab boxes in the command tab.

## Example

{% code-snippet { file-name: CommandsManager.SpacerAndSubMenu.* } %}

The above commands configuration would result in the following menu and command tab boxes created:

![Sub-menus and spacer](sub-menu-and-spacer.png)

* Command1 and Command2 are commands of the top level menu defined in Commands_e enumeration
* Spacer is added between Command1 and Command2
* SubCommand1 and SubCommand2 are commands of SubCommands_e enumeration which is a sub menu of Commands_e enumeration

![Command tab boxes](command-tab.png)

* All commands (including sub menu commands) are added on the same command tab
* Command1 and Command2 are placed in a separate command tab boxes of SubCommand1 and SubCommand2
* Spacer between Command1 and Command2 is ignored in the commands tab
