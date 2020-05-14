---
layout: article
title: Modules in Visual Basic
caption: Modules
description: Article explain the usage of modules for storing the shareable functions and variables in Visual Basic
image: /visual-basic/modules/add-new-module.png
order: 3
---
Modules are containers to define custom functions, procedures or variables to group code in Visual Basic.

Module containing an entry point subroutine (main) is an entry module. It is always at least one module defined in the Visual Basic macro.

In order to add new module it is required to RMB (right mouse button click) the **Modules** folder and select *Inset->Module* command

{% include img.html src="add-new-module.png" height=250 alt="Adding new module to the macro" align="center" %}

Module must have an unique name which can be defined by the developer.

{% include img.html src="module-properties.png" alt="Module properties" align="center" %}

Functions defined in module are public. Members (variables) declared with **Dim** keyword are only visible within this module scope and not visible for another modules, while members declared with **Public** keyword are visible for this and other modules. Refer [Variables Scope](visual-basic/variables/scope) article for more information.

{% include img.html src="module-members.png" alt="Module members" align="center" %}

Module members are available in IntelliSense after typing the name of the module followed by . symbol.

{% include img.html src="module-members-intellisense.png" alt="IntelliSense for members defined in the module" align="center" %}
