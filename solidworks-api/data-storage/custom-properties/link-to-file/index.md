---
layout: sw-tool
caption: Link Custom Properties To File
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

Set the value of the **UPDATE_ON_CSV_FILE_CHANGE_ONLY** constant to **True** or **False** to configure if properties need to reload only if properties text file is changed or always when macro.
 
~~~ vb
Const UPDATE_ON_CSV_FILE_CHANGE_ONLY As Boolean = False
~~~

Macro will ask for the following input parameters upon insertion of the macro feature:

* Should the properties be configuration specific or file specific
* Should the properties be cleared upon update
* Should the reference components of the assembly be included into the scope of the properties

Properties will be automatically updated upon rebuild of the macro feature.

{% code-snippet { file-name: MacroFeature.vba } %}
