---
layout: sw-tool
title: Rename cut list features based on custom properties using SOLIDWORKS API
caption: Rename Cut List Features
description: VBA macro to rename cut list features (sheet metal and weldment) based on custom properties using SOLIDWORKS API
lang: en
image: /solidworks-api/document/cut-lists/rename-cut-list-items/sheet-metal-cut-list.png
labels: [cut-list,sheet metal,weldment,rename]
categories: sw-tools
group: Part
---
{% include img.html src="sheet-metal-cut-list.png" width=250 alt="Sheet metal cut list features" align="center" %}

This VBA macro allows to rename all cut list features for weldment and sheet metal part based on the name template which can include values of custom properties and free text.

{% include img.html src="cut-list-properties.png" width=550 alt="Cut list properties" align="center" %}

To configure the macro modify the values of *NAME_TEMPLATE* and *PROPERTIES* variables

* *NAME_TEMPLATE* can contain free text and 0-based placeholders which will be dynamically replaced by corresponding custom properties values
* Set the names of the properties to extract by assigning the Array of *PROPERTIES* variable in *Init* function

~~~ vb
Const NAME_TEMPLATE = "FreeText_{0}_{1}_{2}_{3}" 'Each feature is renamed with FreeText_ followed by the value of the first custom property specified in PROPERTIES, then _ etc.
Dim PROPERTIES As Variant

Dim swApp As SldWorks.SldWorks

Sub Init(Optional dummy As Variant = Empty)
    PROPERTIES = Array("Prp1", "Prp2", "Prp3", "Prp4") 'custom properties to extract. Value of Prp1 will replace {0}, Prp2 will replace {1} etc.
End Sub
~~~

{% include_relative Macro.vba.codesnippet %}
