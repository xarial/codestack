---
layout: sw-tool
title: Macro to remove all colors from SOLIDWORKS part
caption: Remove All Colors From Part
description: Macro demonstrates how to remove all colors from the part document on all levels (face, feature, body) using SOLIDWORKS API
lang: en
image: /solidworks-api/document/appearance/remove-color/material-properties-levels.png
logo: /solidworks-api/document/appearance/remove-color/remove-colors.svg
labels: [remove color, appearance, material property]
categories: sw-tools
group: Part
---
{% include img.html src="material-properties-levels.png" width=250 alt="Appearance layers in Part document" align="center" %}

This macro removes all colors from the part document on all levels (face, feature, body) using SOLIDWORKS API.

Macro can be configured to remove the colors from all configurations or active configuration only. This option can be set by changing the value of the following constant at the beginning of the macro:

~~~ vb
Const REMOVE_FROM_ALL_CONFIGS As Boolean = True 'True to remove from all configurations, False to remove from active configuration only
~~~

{% include_relative Macro.vba.codesnippet %}
