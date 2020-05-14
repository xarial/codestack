---
layout: article
title: Read configuration specific cut-list property from the selected component using SOLIDWORKS API
caption: Read Component Cut-List Properties
description: VBA macro to read all properties from the cut-list of the selected component in the assembly with respect to the component configuration using SOLIDWORKS API
image: /solidworks-api/data-storage/custom-properties/read-component-cutlist/cut-list-properties.png
labels: [cut-list property,component]
---
{% include img.html src="cut-list-properties.png" width=550 alt="Cut list properties" align="center" %}

This VBA macro example demonstrates how to read and print all custom properties from all cut-list elements of the selected component in assembly using SOLIDWORKS API.

Cut-lists are read from the respective referenced configuration of the component.

Result is output to the immediate window of VBA editor in the following format.

~~~
CS-02-1 (A)
    Sheet<1>
        Bounding Box Length: 150
        Bounding Box Width: 50
        Sheet Metal Thickness: 0.74
        Bounding Box Area: 7500
        Bounding Box Area-Blank: 7500
        Cutting Length-Outer: 400
        Cutting Length-Inner: 0
        Cut Outs: 0
        Bends: 0
        Bend Allowance: 0.5
        Material: Material <not specified>
        Mass: 5.52
        Description: Sheet
        Bend Radius: 0.74
        Surface Treatment: Finish <not specified>
        Cost-TotalCost: 0
        QUANTITY: 1
~~~

{% include_relative Macro.vba.codesnippet %}
