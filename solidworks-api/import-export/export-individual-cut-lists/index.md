---
caption: Export Individual Cut-Lists
title: Export individual bodies from cut-lists from SOLIDWORKS part file via Macro+ framework
description: VBA macro demonstrates how to use Macro+ to export individual bodies from cut-lists to foreign format from the active SOLIDWORKS part
image:
macro-plus: vba
---

This VBA macro is [Macro+](https://cadplus.xarial.com/macro-plus/) enabled macro that allows exporting unique bodies from all cut-lists in the active part file as individual files to foreign format (e.g. STEP, IGES, Parasolid etc.).

This macro supports several custom variables:

* **cutListPrp** with argument for the property name will be resolved to the corresponding cut-list custom property value
* **cutListName** will be resolved to the name of the cut-list item
* **cutListIndex** will be resolved to the cut-list item position in the tree (item number)

> Macro will replace all illegal characters in the file name with underscore

{% code-snippet { file-name: Macro.vba } %}

## CustomVariableValueProvider Class Module

{% code-snippet { file-name: CustomVariableValueProvider.vba } %}