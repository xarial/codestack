---
layout: article
title: VSTA Macro which activates feature manager tab via SOLIDWORKS API
caption: Activate Feature Manager Tab
description: Example demonstrates how to activate standard tabs (feature manager tree, property manager, configuration manager, DimXpert manager, display manager) in the feature manager view using SOLIDWORKS API
lang: en
image: /solidworks-api/document/features-manager/activate-tabs/feature-manager-tabs.png
labels: [feature manager, tab]
---
{% include img.html src="feature-manager-tabs.png" alt="Feature Manager Tabs" align="center" %}

This example demonstrates how to activate standard tabs (feature manager tree, property manager, configuration manager, DimXpert manager, display manager) in the feature manager view using SOLIDWORKS API.

* Specify the tab to activate using the *FeatMgrTab_e* enumeration
* Run the macro (VSTA3)
* Active tab is shown in the message box
* Specified tab is activated

**ModelDocExtension.cs**
{% include_relative ModelDocExtension.cs.codesnippet %}

**SolidWorksMacro.cs**
{% include_relative SolidWorksMacro.cs.codesnippet %}
