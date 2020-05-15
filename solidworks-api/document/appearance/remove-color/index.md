---
layout: sw-tool
title: Macro to remove all colors from SOLIDWORKS part
caption: Remove All Colors From Part
description: Macro demonstrates how to remove all colors from the part document on all levels (face, feature, body) using SOLIDWORKS API
image: remove-colors.svg
labels: [remove color, appearance, material property]
group: Part
---
![Appearance layers in Part document](material-properties-levels.png){ width=250 }

This macro removes all colors from the part document on all levels (face, feature, body) using SOLIDWORKS API.

Macro can be configured to remove the colors from all configurations or active configuration only. This option can be set by changing the value of the following constant at the beginning of the macro:

~~~ vb
Const REMOVE_FROM_ALL_CONFIGS As Boolean = True 'True to remove from all configurations, False to remove from active configuration only
~~~

{% code-snippet { file-name: Macro.vba } %}
