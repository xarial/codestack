---
layout: article
title: Custom enable command state for SOLIDWORKS commands
caption: Custom Enable Command State
description: Explanation on using the custom enable states for the SOLIDWORKS commands using SwEx framework
toc-group-name: labs-solidworks-swex
order: 4
---
There are 4 command states supported by SOLIDWORKS:

1. Deselected and enabled. This is default option when button can be clicked
1. Deselected and disabled. This option is used when the command is not supported in certain framework. For example mate command will be disabled in parts and drawings as it is only supported in the assemblies.
1. Selected and disabled. This represents the disabled checked button
1. Selected and enabled. This represents checked button

![Supported command states](command-states.png)

SwEx framework will assign the appropriate state (enabled or disabled) for the commands based on their supported workspaces if defined in the [CommandItemInfoAttribute](https://docs.codestack.net/swex/add-in/html/T_CodeStack_SwEx_AddIn_Attributes_CommandItemInfoAttribute.htm). However user can alter the state to provide more advanced management (for example it might be required to enable command if certain object is selected or if any bodies or components are present in the model). To do this it is required to specify the enable method handler as the last parameter of [AddCommandGroup](https://docs.codestack.net/swex/add-in/html/M_CodeStack_SwEx_AddIn_SwAddInEx_AddCommandGroup__1.htm) or [AddContextMenu](https://docs.codestack.net/swex/add-in/html/M_CodeStack_SwEx_AddIn_SwAddInEx_AddContextMenu__1.htm) methods.

The method is defined as [EnableMethodDelegate](https://docs.codestack.net/swex/add-in/html/T_CodeStack_SwEx_AddIn_EnableMethodDelegate_1.htm) delegate and provides the command id as first parameter and state passed by reference as second parameter.

The value of state will be preassigned based on the workspace and can be changed by the user within the method.

> This method allows to implement the toggle button in toolbar and menu. To set the checked state assign the *SelectEnable* or *SelectDisable* values of [CommandItemEnableState_e](https://docs.codestack.net/swex/add-in/html/T_CodeStack_SwEx_AddIn_Enums_CommandItemEnableState_e.htm) enumeration.

{#% include code-tabs.html src="CommandsManager.CustomEnableState" %}