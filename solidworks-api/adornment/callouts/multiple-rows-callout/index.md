---
title: Create multiple rows callout using SOLIDWORKS API
caption: Create Multiple Rows Callout
description: Example demonstrates how to create a callout with multiple rows from the selection in SOLIDWORKS API
image: sw-callout-spec.png
labels: [adornment, callout, example, note, solidworks api]
redirect-from:
  - /2018/04/solidworks-api-adornment-create-multirow-callout.html
---
This example demonstrates how to create a callout with multiple rows while selecting the object using [ISelectionMgr::CreateCallout2](http://help.solidworks.com/2018/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.iselectionmgr~createcallout2.html) SOLIDWORKS API method.

![Callout element specification](sw-callout-spec.png){ width=640 height=354 }

First row of the displayed callout is not editable (read only). Value of second row can be changed. The changed value will be displayed in the message box.

*Macro*

{% code-snippet { file-name: Macro.vba } %}

*CalloutHandler class*

{% code-snippet { file-name: CalloutHandler.vba } %}
