---
layout: sw-tool
caption: Link Custom Properties To Fie
title: Link SOLIDWORKS custom properties from text file
description: VBA macro to link and auto-update multiple SOLIDWORKS custom properties from the external CSV/Text file into configuration or file
image: link-custom-property-file.svg
group: Custom Properties
---
This VBA macro allows to link external comma separated file into the configuration specific or file specific custom properties of SOLIDWORKS file.

CSV file consists of 2 columns (property name and property value) without a header.

If value of the cell contains special symbol **"** then the cell must have **""** at the start and ant the end of the cell value.

~~~
Company,Xarial Pty Limited
Material,"""SW-Material"""
Mass,"""SW-Mass"""
~~~

> You can use Excel to modify these values and export to CSV file with comma delimiter, special symbol will be formatted correctly automatically.

> Commas and new line symbols in the property names or values are not supported

Set the value of the **CLEAR_PROPERTIES** constant to **True** or **False** to configure if existing properties need to be deleted before updating.

~~~ vb
Const CLEAR_PROPERTIES As Boolean = False
~~~

{% code-snippet { file-name: Macro.vba } %}

In order to dynamically link external text file and update properties on every rebuild, use the following macro.

{% code-snippet { file-name: MacroFeature.vba } %}