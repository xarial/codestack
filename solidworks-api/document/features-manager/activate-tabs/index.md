---
title: VSTA Macro which activates feature manager tab via SOLIDWORKS API
caption: Activate Feature Manager Tab
description: Example demonstrates how to activate standard tabs (feature manager tree, property manager, configuration manager, DimXpert manager, display manager) in the feature manager view using SOLIDWORKS API
image: feature-manager-tabs.png
labels: [feature manager, tab]
---
![Feature Manager Tabs](feature-manager-tabs.png)

This example demonstrates how to activate standard tabs (feature manager tree, property manager, configuration manager, DimXpert manager, display manager) in the feature manager view using SOLIDWORKS API.

* Specify the tab to activate using the *FeatMgrTab_e* enumeration
* Run the macro (VSTA3)
* Active tab is shown in the message box
* Specified tab is activated

**ModelDocExtension.cs**
{% code-snippet { file-name: ModelDocExtension.cs } %}

**SolidWorksMacro.cs**
{% code-snippet { file-name: SolidWorksMacro.cs } %}
