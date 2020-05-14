---
layout: article
title: Modify the definition of axis feature using SOLIDWORKS API
caption: Modify Axis Definition
description: VBA macro example to change the selection of the axis feature using SOLIDWORKS API
image: /solidworks-api/document/features-manager/modify-axis-definition/axis-definition.png
labels: [axis,definition]
---
{% include img.html src="axis-definition.png" width=250 alt="Axis property manager page" align="center" %}

This VBA example demonstrates how to modify the definition and change the selection for the axis feature using SOLIDWORKS API.

* Select the target axis feature to modify as first selection
* Select the objects to be set as the reference for the target axis. For example 2 intersecting planes, edge, etc.

As the result the selected objects (second to the last) will be assigned to the axis (first selection).

{% code-snippet { file-name: Macro.vba } %}
