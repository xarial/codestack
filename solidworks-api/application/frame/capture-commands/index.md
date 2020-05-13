---
layout: sw-tool
title: VBA macro to capture SOLIDWORKS commands via API event handlers
caption: Capture SOLIDWORKS Commands
description: Macro allows capturing SOLIDWORKS and user commands into the list box
lang: en
image: /solidworks-api/application/frame/capture-commands/capturing-hide-command-id.png
labels: [command, event]
categories: sw-tools
group: Developers
---
This macro allows capturing of SOLIDWORKS command ids (e.g. toolbar, page button or context menu clicks). Commands are defined in the [swCommands_e](http://help.solidworks.com/2012/english/api/swcommands/solidworks.interop.swcommands~solidworks.interop.swcommands.swcommands_e.html) enumeration and can be called using the [ISldWorks::RunCommand](http://help.solidworks.com/2012/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.isldworks~runcommand.html) SOLIDWORKS API method.

This could be in particularly useful when certain SOLIDWORKS APIs are not available in the SDK.

All commands have user friendly names however they could not always match the names in the user interface. This fact could make it hard to find the correct command (as there are currently more than 3000 commands available). For example Hide Sketch command in User Interface corresponds to *swCommands_Blank_Refgeom* command id.

### Capturing standard commands

This macro helps to capture the id of command directly from SOLIDWORKS by clicking the required command.

* Run the macro. Form with list is displayed
* Perform the required action (i.e. click button or menu item)
* Command id is recorded and displayed in the list

{% include img.html src="capturing-hide-command-id.png" width=350 alt="Capturing sketch hide command id" align="center" %}

The command id can be looked up in the the [commands list]((http://help.solidworks.com/2012/english/api/swcommands/solidworks.interop.swcommands~solidworks.interop.swcommands.swcommands_e.html))

{% include img.html src="sw-commands-id.png" width=350 alt="Hide sketch command id in swCommands_e enumeration" align="center" %}

> It is not required to explicitly use [swCommands_e](http://help.solidworks.com/2012/english/api/swcommands/solidworks.interop.swcommands~solidworks.interop.swcommands.swcommands_e.html) enumeration as it is defined in another interop (*solidworks.interop.swcommands.dll*). Instead command id can be defined as an integer or custom enumeration.

### Capturing commands from the custom add-ins

For the standard SOLIDWORKS commands, User Command argument will be equal to 0. However commands cannot be defined for any custom add-in or [Macro Buttons]({{ "solidworks-api/getting-started/macros/macro-buttons/" | relative_url }})

If this command is clicked, the command id would be equal to one of the following:

{% include img.html src="user-commands.png" width=450 alt="User specific command ids" align="center" %}

Command would indicate the type of the button (minimized toolbar, menu, macro button etc.), and the User Command Id will be equal to the user id of a custom button. This is a command user id which can be retrieved via [ICommandGroup::CommandId](http://help.solidworks.com/2012/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.ICommandGroup~CommandID.html) property while creating the custom commands manager in the SOLIDWORKS add-in.

{% include img.html src="capturing-user-command-id.png" width=250 alt="Capturing the commands from the custom add-in" align="center" %}

### Creating macro

* Add User Form module to the macro and name it *CommandsMonitorForm*

{% include img.html src="vba-macro-project.png" width=450 alt="VBA project structure" align="center" %}

* Drag-n-drop the List Box control onto the form and name it *lstLog*

{% include img.html src="add-list-box-control.png" width=450 alt="Adding list box control to the form" align="center" %}

* Add the code to corresponding modules

**Macro**

{% include_relative Macro.vba.codesnippet %}

**CommandsMonitorForm**

{% include_relative CommandsMonitorForm.vba.codesnippet %}
