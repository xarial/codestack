---
layout: sw-tool
title: SOLIDWORKS Macro to display callouts with diameters for edges
caption: Display Callouts With Diameters For All Selected Circular Edges
description: Macro will display the callouts with the diameter values of all selected circular edges in the 3D model
image: /solidworks-api/adornment/callouts/circular-edges-display-callouts/edge-callout.png
logo: /solidworks-api/adornment/callouts/circular-edges-display-callouts/edge-callout.svg
labels: [adornment, callout, diameter, edge, example, macro, solidworks api, unit conversion]
categories: sw-tools
group: Model
redirect_from:
  - /2018/03/display-callouts-with-diameters-for-all.html
---
This macro will display the callouts with the diameter values of all selected circular edges in the 3D model using [ISelectionMgr::CreateCallout2](http://help.solidworks.com/2018/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.iselectionmgr~createcallout2.html) SOLIDWORKS API method.

This can be useful while inspecting the model and it is required to see multiple diameters at the same time.

{% include img.html src="hole-diams.png" width=400 height=290 alt="Diameters displayed in the callout for selected holes" align="center" %}

The callout is a visual element in SOLIDWORKS which displays data organized in key-value pairs (single or multiple rows). The callout elements are used in some standard SOLIDWORKS tools such as [Measure tool](http://help.solidworks.com/2017/english/solidworks/sldworks/t_using_the_measure_tool.htm). Usually callouts are attached to the selection and destroyed once the object is deselected.

To run the macro:

1. Select circular edges and run the macro
1. Callouts with the diameter value in the model's units are displayed for all circular edges
1. Clear the selection to remove the callouts

Create new macro and copy the following code into the macro's module:

{% include img.html src="macro-module.png" width=640 height=230 alt="Macro module in VBA editor" align="center" %}

{% code-snippet { file-name: Macro.vba } %}

Create new class module and name it *HoleDiamCalloutHandler.*  

{% include img.html src="insert-class-module.png" width=320 height=220 alt="Adding class module to VBA macro" align="center" %}

Copy the following code in there:

{% code-snippet { file-name: HoleDiamCalloutHandler.vba } %}
